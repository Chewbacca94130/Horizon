"""Contenu déposé dans une note WYSIWYG : fichiers, courriels Outlook, URLs.

Deux temps, imposés par le glisser-déposer Windows :

1. `parse_drop(mime)` lit le `QMimeData` **pendant** l'événement de dépôt — il n'est
   valide que jusque-là — et le réduit à des objets Python simples.
2. `DropImporter.import_items()` archive ensuite le contenu et décrit les blocs à
   insérer. Il tourne *après* le retour du dépôt : Outlook est encore dans sa boucle
   DoDragDrop tant que le dépôt n'est pas terminé, et l'interroger par COM à ce
   moment-là peut bloquer les deux applications.

Les blocs sont des dictionnaires prêts pour le plugin `HorizonAttachment` de
`web/src/wysiwyg.js` (clés en camelCase).
"""

import email
import email.policy
import html as _html
import re
import ssl
import struct
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlsplit

import certifi
import httpx

from horizon.attachment import AttachmentRepository, guess_mime_type, safe_file_name
from horizon.attachment_preview import pdf_page_count
from horizon.document import extract_document_link
from horizon.image import ImageRepository
from horizon.note import html_to_text

_WIN_MIME = 'application/x-qt-windows-mime;value="{}"'
_FILE_GROUP_DESCRIPTOR_W = _WIN_MIME.format("FileGroupDescriptorW")
_FILE_CONTENTS = _WIN_MIME.format("FileContents")
_OUTLOOK_MESSAGES = _WIN_MIME.format("RenPrivateMessages")
_MOZ_URL = "text/x-moz-url"

# Les formats d'image que la note sait afficher (voir webeditor._image_mime) : déposés,
# ils deviennent des images de la note et non des pièces jointes.
_INLINE_IMAGE_SUFFIXES = frozenset({".png", ".jpg", ".jpeg", ".gif", ".webp"})

_EXCERPT_MAX = 2000  # caractères du corps d'un courriel gardés dans la note
_TITLE_MAX = 300
_TITLE_FETCH_BYTES = 256 * 1024
_TITLE_RE = re.compile(r"<title[^>]*>(.*?)</title>", re.IGNORECASE | re.DOTALL)

_OL_EXPLORER = 34  # OlObjectClass.olExplorer
_OL_INSPECTOR = 35  # OlObjectClass.olInspector
_OL_MSG_UNICODE = 9  # OlSaveAsType.olMSGUnicode


class DropImportError(Exception):
    """Contenu déposé qui ne peut pas être archivé ; le message s'adresse à l'utilisateur."""


# ── Ce que le dépôt contenait ─────────────────────────────────────────────────


@dataclass
class DroppedFile:
    path: Path


@dataclass
class DroppedLink:
    url: str
    title: str  # vide : à déduire de l'URL, puis à chercher sur la page


@dataclass
class DroppedOutlookItems:
    """Éléments « virtuels » d'Outlook : courriels (`.msg`) ou pièces jointes d'un
    courriel. Ils n'existent pas sur le disque ; seuls leurs noms sont annoncés."""

    names: list[str]
    contents: bytes = b""  # FileContents, quand Qt a pu le lire (un seul fichier)


DroppedItem = DroppedFile | DroppedLink | DroppedOutlookItems


def _utf16(raw: bytes) -> str:
    return raw.decode("utf-16-le", errors="ignore").split("\x00", 1)[0]


def _file_group_names(raw: bytes) -> list[str]:
    """Noms d'un FILEGROUPDESCRIPTORW : un compteur, puis des FILEDESCRIPTORW de 592
    octets dont le nom (260 WCHAR) commence à l'octet 72."""
    if len(raw) < 4:
        return []
    (count,) = struct.unpack_from("<I", raw, 0)
    names = []
    for i in range(min(count, 256)):
        start = 4 + i * 592 + 72
        chunk = raw[start:start + 520]
        if len(chunk) < 2:
            break
        name = _utf16(chunk).strip()
        if name:
            names.append(name)
    return names


def _is_web_url(url: str) -> bool:
    return url.startswith(("http://", "https://"))


def parse_drop(mime, read_contents: bool = True) -> list[DroppedItem] | None:
    """Ce que la note doit archiver, ou None quand le dépôt ne la concerne pas — texte
    ou HTML glissé depuis une page, image d'une page web : l'éditeur les colle alors
    lui-même, comme avant."""
    if mime.hasFormat(_FILE_GROUP_DESCRIPTOR_W) and not mime.hasUrls():
        names = _file_group_names(bytes(mime.data(_FILE_GROUP_DESCRIPTOR_W).data()))
        if names:
            contents = b""
            if read_contents and len(names) == 1 and mime.hasFormat(_FILE_CONTENTS):
                contents = bytes(mime.data(_FILE_CONTENTS).data())
            return [DroppedOutlookItems(names, contents)]

    if mime.hasUrls():
        urls = mime.urls()
        files = [Path(u.toLocalFile()) for u in urls if u.isLocalFile()]
        if files:
            return [DroppedFile(p) for p in files]
        web = [u.toString() for u in urls if _is_web_url(u.toString())]
        if web:
            # Une image glissée depuis une page : l'éditeur la gère (balise <img>).
            html = mime.html() if mime.hasHtml() else ""
            if "<img" in html.lower() and "<a" not in html.lower():
                return None
            return [DroppedLink(url, _link_title(mime, url)) for url in web]

    # SharePoint et certaines pages ne publient le lien qu'en HTML ou en texte.
    if mime.hasText() and not mime.hasHtml():
        text = mime.text().strip()
        if _is_web_url(text) and len(text.split()) == 1:
            return [DroppedLink(text, "")]
    return None


def _link_title(mime, url: str) -> str:
    """Titre fourni par le navigateur : text/x-moz-url (« url\ntitre ») ou le texte de
    l'ancre HTML qui pointe vers l'URL. Vide s'il n'y en a pas ou s'il répète l'URL."""
    title = ""
    if mime.hasFormat(_MOZ_URL):
        lines = _utf16(bytes(mime.data(_MOZ_URL).data())).splitlines()
        if len(lines) >= 2 and lines[0].strip() == url:
            title = lines[1].strip()
    if not title and mime.hasHtml():
        found = extract_document_link(mime.html(), None)
        if found is not None and found[1] == url:
            title = found[0]
    return "" if title == url else title


# ── Archivage ─────────────────────────────────────────────────────────────────


@dataclass
class ImportResult:
    blocks: list[dict] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    titles_to_fetch: list[tuple[str, str]] = field(default_factory=list)  # (url, titre provisoire)


def host_title(url: str) -> str:
    host = urlsplit(url).hostname or url
    return host.removeprefix("www.")


def _excerpt(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t ]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if len(text) > _EXCERPT_MAX:
        text = text[:_EXCERPT_MAX].rstrip() + " …"
    return text


def _format_date(value) -> str:
    if value is None:
        return ""
    try:
        return value.strftime("%d/%m/%Y %H:%M")
    except (AttributeError, ValueError):
        return str(value)


class DropImporter:
    def __init__(self, attachment_repo: AttachmentRepository, image_repo: ImageRepository | None) -> None:
        self._attachment_repo = attachment_repo
        self._image_repo = image_repo

    def import_items(self, items: list[DroppedItem], note_id: int | None) -> ImportResult:
        result = ImportResult()
        for item in items:
            try:
                if isinstance(item, DroppedFile):
                    result.blocks.append(self._import_file(item.path, note_id))
                elif isinstance(item, DroppedLink):
                    title = item.title.strip()[:_TITLE_MAX]
                    if not title:
                        title = host_title(item.url)
                        result.titles_to_fetch.append((item.url, title))
                    result.blocks.append({"kind": "link", "href": item.url, "title": title})
                else:
                    result.blocks.extend(self._import_outlook(item, note_id))
            except DropImportError as exc:
                result.errors.append(str(exc))
            except Exception as exc:  # un élément illisible ne doit pas perdre les autres
                print(f"Horizon: drop import failed: {exc!r}", file=sys.stderr)
                result.errors.append(f"Import impossible : {exc}")
        return result

    # ── Fichiers ──────────────────────────────────────────────────────────
    def _import_file(self, path: Path, note_id: int | None) -> dict:
        if path.is_dir():
            raise DropImportError(f"« {path.name} » est un dossier : déposez les fichiers qu'il contient.")
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise DropImportError(f"Lecture de « {path.name} » impossible : {exc.strerror or exc}") from exc
        return self._block_for_bytes(path.name, data, note_id, source=path)

    def _block_for_bytes(self, name: str, data: bytes, note_id: int | None, source: Path | None = None) -> dict:
        suffix = Path(name).suffix.lower()
        if suffix in _INLINE_IMAGE_SUFFIXES and self._image_repo is not None:
            return {"kind": "image", "imageId": self._image_repo.create(data)}

        if suffix == ".eml":
            meta = _eml_metadata(data)
            return self._email_block(name, data, note_id, meta)
        if suffix == ".msg" and source is not None:
            meta = _msg_metadata_via_outlook(source)
            if meta is not None:
                return self._email_block(name, data, note_id, meta)

        attachment = self._create(name, data, note_id)
        pages = pdf_page_count(data) if suffix == ".pdf" or data[:5] == b"%PDF-" else 0
        return {
            "kind": "file",
            "attachmentId": attachment.id,
            "name": attachment.name,
            "size": attachment.size,
            "mime": attachment.mime_type,
            "pages": pages,
            "display": "preview" if pages > 0 else "icon",
        }

    def _create(self, name: str, data: bytes, note_id: int | None):
        try:
            return self._attachment_repo.create(name, data, note_id=note_id, mime_type=guess_mime_type(name))
        except ValueError as exc:
            raise DropImportError(str(exc)) from exc

    def _email_block(self, name: str, data: bytes, note_id: int | None, meta: dict) -> dict:
        attachment = self._create(name, data, note_id)
        return {
            "kind": "email",
            "attachmentId": attachment.id,
            "name": attachment.name,
            "size": attachment.size,
            "mime": attachment.mime_type,
            "subject": meta.get("subject") or Path(name).stem,
            "sender": meta.get("sender", ""),
            "date": meta.get("date", ""),
            "excerpt": _excerpt(meta.get("body", "")),
            "display": "preview",
        }

    # ── Outlook ───────────────────────────────────────────────────────────
    def _import_outlook(self, item: DroppedOutlookItems, note_id: int | None) -> list[dict]:
        # Une pièce jointe seule dont Qt a pu lire le contenu : pas besoin d'Outlook.
        if item.contents and len(item.names) == 1 and not item.names[0].lower().endswith(".msg"):
            return [self._block_for_bytes(item.names[0], item.contents, note_id)]

        blocks = []
        for exported in export_outlook_items(item.names):
            if exported.meta is not None:
                blocks.append(self._email_block(exported.name, exported.data, note_id, exported.meta))
            else:
                blocks.append(self._block_for_bytes(exported.name, exported.data, note_id))
        if not blocks:
            raise DropImportError(
                "Outlook n'a renvoyé aucun élément : sélectionnez le courriel dans Outlook puis glissez-le à nouveau."
            )
        return blocks


# ── Courriels ─────────────────────────────────────────────────────────────────


def _eml_metadata(data: bytes) -> dict:
    message = email.message_from_bytes(data, policy=email.policy.default)
    body = ""
    part = message.get_body(preferencelist=("plain", "html"))
    if part is not None:
        try:
            content = part.get_content()
        except (LookupError, ValueError):
            content = ""
        body = html_to_text(content) if part.get_content_type() == "text/html" else content
    date = ""
    if message["date"] is not None:
        try:
            date = _format_date(message["date"].datetime.astimezone())
        except (AttributeError, ValueError, TypeError):
            date = str(message["date"])
    return {
        "subject": str(message["subject"] or ""),
        "sender": str(message["from"] or ""),
        "date": date,
        "body": body,
    }


def _outlook_application():
    """Outlook *classique* déjà ouvert, ou DropImportError. Jamais lancé : un dépôt ne
    doit pas démarrer Outlook (et le nouvel Outlook n'a pas d'interface COM)."""
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError as exc:
        raise DropImportError("pywin32 n'est pas installé : import Outlook indisponible.") from exc
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except pywintypes.com_error as exc:
        raise DropImportError(
            "Outlook classique n'est pas accessible : il doit être ouvert pour importer un courriel."
        ) from exc


def _item_metadata(item) -> dict | None:
    """Métadonnées d'un élément Outlook (courriel, invitation…), None s'il n'en a pas."""
    subject = getattr(item, "Subject", None)
    if subject is None:
        return None
    sender = getattr(item, "SenderName", "") or ""
    address = getattr(item, "SenderEmailAddress", "") or ""
    if "@" in address and address.lower() not in sender.lower():
        sender = f"{sender} <{address}>" if sender else address
    date = None
    for attribute in ("ReceivedTime", "SentOn", "Start", "CreationTime"):
        date = getattr(item, attribute, None)
        if date is not None:
            break
    return {
        "subject": str(subject),
        "sender": sender,
        "date": _format_date(date),
        "body": str(getattr(item, "Body", "") or ""),
    }


def _msg_metadata_via_outlook(path: Path) -> dict | None:
    """Métadonnées d'un .msg sur le disque, si Outlook est ouvert ; sinon le fichier
    reste une pièce jointe ordinaire."""
    try:
        outlook = _outlook_application()
        item = outlook.Session.OpenSharedItem(str(path))
    except Exception:
        return None
    try:
        return _item_metadata(item)
    finally:
        try:
            item.Close(1)  # olDiscard
        except Exception:
            pass


@dataclass
class OutlookExport:
    name: str
    data: bytes
    meta: dict | None  # None : pièce jointe d'un courriel, pas un courriel


def _selected_items(window) -> list:
    if window.Class == _OL_INSPECTOR:
        return [window.CurrentItem]
    selection = window.Selection
    return [selection.Item(i) for i in range(1, selection.Count + 1)]


def _selected_attachments(window) -> list:
    try:
        selection = window.AttachmentSelection
    except Exception:
        return []
    return [selection.Item(i) for i in range(1, selection.Count + 1)]


def export_outlook_items(names: list[str]) -> list[OutlookExport]:
    """Récupère par COM les éléments glissés depuis Outlook.

    Le glisser d'Outlook n'annonce que des noms (« Objet du courriel.msg »,
    « Rapport.pdf ») : leur contenu est lu dans la sélection de la fenêtre active
    d'Outlook, qui est forcément celle d'où vient le glisser."""
    outlook = _outlook_application()
    window = outlook.ActiveWindow()
    if window is None or window.Class not in (_OL_EXPLORER, _OL_INSPECTOR):
        raise DropImportError("Aucune fenêtre Outlook active : glissez le courriel depuis Outlook.")

    wanted_files = {n.lower() for n in names if not n.lower().endswith(".msg")}
    folder = Path(tempfile.mkdtemp(prefix="horizon-outlook-"))
    exports: list[OutlookExport] = []
    try:
        if wanted_files:
            for attachment in _selected_attachments(window):
                file_name = str(attachment.FileName)
                if file_name.lower() not in wanted_files:
                    continue
                path = folder / safe_file_name(file_name)
                attachment.SaveAsFile(str(path))
                exports.append(OutlookExport(file_name, path.read_bytes(), None))
        if len(exports) < len(names):
            for item in _selected_items(window):
                meta = _item_metadata(item)
                if meta is None:
                    continue
                name = safe_file_name(meta["subject"] or "Courriel") + ".msg"
                path = folder / f"{len(exports)}-{name}"
                item.SaveAs(str(path), _OL_MSG_UNICODE)
                exports.append(OutlookExport(name, path.read_bytes(), meta))
    finally:
        for child in folder.iterdir():
            child.unlink(missing_ok=True)
        folder.rmdir()
    return exports


# ── Titre d'une page ──────────────────────────────────────────────────────────


def fetch_page_title(url: str, timeout: float = 6.0) -> str | None:
    """`<title>` d'une page web, ou None (hors ligne, page protégée, pas de titre).

    Appelé hors du thread de l'interface : seule la page déposée est interrogée."""
    context = ssl.create_default_context(cafile=certifi.where())
    try:
        with httpx.Client(
            verify=context, timeout=timeout, follow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Horizon"},
        ) as client:
            with client.stream("GET", url) as response:
                if response.status_code >= 400:
                    return None
                if "html" not in response.headers.get("content-type", "html"):
                    return None
                raw = b""
                for chunk in response.iter_bytes():
                    raw += chunk
                    if len(raw) >= _TITLE_FETCH_BYTES or b"</title>" in raw.lower():
                        break
                encoding = response.encoding or "utf-8"
    except (httpx.HTTPError, ssl.SSLError, OSError, ValueError):
        return None
    match = _TITLE_RE.search(raw.decode(encoding, errors="replace"))
    if match is None:
        return None
    title = " ".join(_html.unescape(match.group(1)).split())
    return title[:_TITLE_MAX] or None

