from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class Attachment:
    name: str
    size: int                  # actual size in bytes from MAPI
    data: bytes                # raw bytes; empty if skipped or embedded msg
    is_embedded_msg: bool = False
    skipped: bool = False      # True if size exceeded the limit


@dataclass
class Email:
    folder_path: str
    subject: str
    from_: str
    to: list[str]
    cc: list[str]
    date: Optional[datetime]
    message_id: str
    body_text: str
    body_html: str
    attachments: list[Attachment] = field(default_factory=list)
