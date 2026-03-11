"""PST file traversal using pypff."""

from __future__ import annotations

import email as email_lib
import logging
import re
import struct
from datetime import datetime, timezone
from typing import Iterator

import pypff

from models import Attachment, Email

logger = logging.getLogger(__name__)

# MAPI property entry types
_PR_MESSAGE_CLASS       = 0x001A
_PR_DISPLAY_TO          = 0x0E04
_PR_DISPLAY_CC          = 0x0E03
_PR_SENDER_NAME         = 0x0C1A
_PR_SENDER_EMAIL        = 0x0C1F
_PR_DISPLAY_NAME        = 0x3001
_PR_ATTACH_METHOD       = 0x3705
_PR_ATTACH_LONG_FNAME   = 0x3707
_PR_ATTACH_FNAME        = 0x3704

_ATTACH_METHOD_EMBEDDED = 5
_EMAIL_CLASS_PREFIX     = "IPM.Note"


def _safe_str(value, encoding="utf-8") -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode(encoding, errors="replace")
    return str(value)


def _read_mapi_props(obj) -> tuple[dict[int, str], dict[int, int]]:
    """Return (string_props, int_props) dicts keyed by MAPI entry_type."""
    strings: dict[int, str] = {}
    ints: dict[int, int] = {}
    try:
        for rs_i in range(obj.number_of_record_sets):
            rs = obj.get_record_set(rs_i)
            for e_i in range(rs.number_of_entries):
                e = rs.get_entry(e_i)
                if not e.data:
                    continue
                vt, et = e.value_type, e.entry_type
                if vt == 0x1F:
                    strings[et] = e.data.decode("utf-16-le", errors="replace").rstrip("\x00")
                elif vt == 0x1E:
                    strings[et] = e.data.decode("latin-1", errors="replace").rstrip("\x00")
                elif vt == 0x03 and len(e.data) == 4:
                    ints[et] = struct.unpack_from("<I", e.data)[0]
    except Exception as exc:
        logger.debug("Could not read MAPI props: %s", exc)
    return strings, ints


def _get_message_class(message) -> str:
    strings, _ = _read_mapi_props(message)
    return strings.get(_PR_MESSAGE_CLASS, "IPM.Note")


def _parse_headers(message) -> dict[str, str]:
    raw = _safe_str(message.transport_headers)
    if not raw:
        return {}
    parsed = email_lib.message_from_string(raw)
    return {k.lower(): v for k, v in parsed.items()}


def _parse_date(message) -> datetime | None:
    try:
        dt = message.delivery_time or message.client_submit_time or message.creation_time
        if dt is None:
            return None
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _parse_attachment(attachment, index: int, max_bytes: int) -> Attachment:
    strings, ints = _read_mapi_props(attachment)

    is_embedded = ints.get(_PR_ATTACH_METHOD) == _ATTACH_METHOD_EMBEDDED

    # Determine display name
    if is_embedded:
        display = strings.get(_PR_DISPLAY_NAME, "")
        name = re.sub(r"\s*\(\d+(\.\d+)?\s*[KMG]?B\)\s*$", "", display, flags=re.IGNORECASE).strip()
        name = (name or f"embedded_msg_{index + 1}") + ".msg"
    else:
        name = strings.get(_PR_ATTACH_LONG_FNAME) or strings.get(_PR_ATTACH_FNAME) or f"attachment_{index + 1}"

    size = attachment.size or 0

    # Embedded messages: don't try to read raw bytes — the stream is a full
    # MAPI message object, not a simple binary blob.
    if is_embedded:
        return Attachment(name=name, size=size, data=b"", is_embedded_msg=True)

    # max_bytes == -1 is a special sentinel meaning "don't read any data, no warning"
    if max_bytes == -1:
        return Attachment(name=name, size=size, data=b"", skipped=True)

    # Size guard — skip reading data if it exceeds the limit
    if max_bytes > 0 and size > max_bytes:
        logger.warning(
            "Attachment '%s' is %s — exceeds limit of %s, skipping embed",
            name, _fmt_size(size), _fmt_size(max_bytes),
        )
        return Attachment(name=name, size=size, data=b"", skipped=True)

    try:
        attachment.seek_offset(0)
        data = attachment.read_buffer(size) if size else b""
    except Exception as exc:
        logger.warning("Could not read attachment data for '%s': %s", name, exc)
        data = b""

    return Attachment(name=name, size=size, data=data)


def _parse_attachments(message, max_bytes: int) -> list[Attachment]:
    result: list[Attachment] = []
    try:
        for i in range(message.number_of_attachments):
            try:
                result.append(_parse_attachment(message.get_attachment(i), i, max_bytes))
            except Exception as exc:
                logger.debug("Could not parse attachment %d: %s", i, exc)
    except Exception as exc:
        logger.debug("Could not enumerate attachments: %s", exc)
    return result


def _split_addrs(value: str) -> list[str]:
    return [a.strip() for a in value.split(";") if a.strip()] if value else []


def _parse_message(message, folder_path: str, max_attachment_bytes: int) -> Email:
    headers = _parse_headers(message)
    strings, _ = _read_mapi_props(message)

    subject = _safe_str(message.subject)

    if "from" in headers:
        from_ = headers["from"]
    else:
        name = strings.get(_PR_SENDER_NAME) or _safe_str(message.sender_name)
        addr = strings.get(_PR_SENDER_EMAIL, "")
        from_ = f"{name} <{addr}>" if name and addr else (name or addr)

    if "to" in headers:
        to_list = [a.strip() for a in headers["to"].split(",") if a.strip()]
    else:
        to_list = _split_addrs(strings.get(_PR_DISPLAY_TO, ""))

    if "cc" in headers:
        cc_list = [a.strip() for a in headers["cc"].split(",") if a.strip()]
    else:
        cc_list = _split_addrs(strings.get(_PR_DISPLAY_CC, ""))

    return Email(
        folder_path=folder_path,
        subject=subject,
        from_=from_,
        to=to_list,
        cc=cc_list,
        date=_parse_date(message),
        message_id=headers.get("message-id", ""),
        body_text=_safe_str(message.plain_text_body),
        body_html=_safe_str(message.html_body),
        attachments=_parse_attachments(message, max_attachment_bytes),
    )


def _walk_folder(folder, path: str, email_only: bool, max_attachment_bytes: int) -> Iterator[Email]:
    folder_name = _safe_str(folder.name) or "Unknown"
    current_path = f"{path}/{folder_name}" if path else folder_name

    try:
        for i in range(folder.number_of_sub_messages):
            try:
                message = folder.get_sub_message(i)
                if email_only:
                    msg_class = _get_message_class(message)
                    if not msg_class.startswith(_EMAIL_CLASS_PREFIX):
                        logger.debug("Skipping non-email item (class=%s) in %s", msg_class, current_path)
                        continue
                yield _parse_message(message, current_path, max_attachment_bytes)
            except Exception as exc:
                logger.warning("Skipping malformed message %d in %s: %s", i, current_path, exc)
    except Exception as exc:
        logger.warning("Could not read messages in folder %s: %s", current_path, exc)

    try:
        for i in range(folder.number_of_sub_folders):
            try:
                subfolder = folder.get_sub_folder(i)
                yield from _walk_folder(subfolder, current_path, email_only, max_attachment_bytes)
            except Exception as exc:
                logger.warning("Skipping subfolder %d in %s: %s", i, current_path, exc)
    except Exception as exc:
        logger.warning("Could not read subfolders in %s: %s", current_path, exc)


def parse_pst(
    pst_path: str,
    email_only: bool = True,
    max_attachment_bytes: int = 10 * 1024 * 1024,
) -> Iterator[Email]:
    """Yield Email objects for every message in the PST file."""
    pst = pypff.file()
    pst.open(pst_path)
    try:
        root = pst.get_root_folder()
        yield from _walk_folder(root, "", email_only, max_attachment_bytes)
    finally:
        pst.close()


def _fmt_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024
