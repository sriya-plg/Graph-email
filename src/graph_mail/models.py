from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class Attachment:
    id: str
    message_id: str
    name: str
    content_type: str
    size: int
    is_inline: bool
    content: bytes

    def save(self, directory: str | Path) -> Path:
        directory = Path(directory)
        directory.mkdir(parents=True, exist_ok=True)
        safe = "".join(c if c not in '\\/:*?"<>|' else "_" for c in self.name) or self.id
        path = directory / safe
        path.write_bytes(self.content)
        return path
