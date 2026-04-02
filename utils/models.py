from datetime import datetime
from pathlib import Path
from typing import Dict
from pptx import Presentation

class PresentationFile:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.last_modified = datetime.now()
        self.pptx_object = Presentation(file_path)

    def _serialize_file_info(self) -> Dict:
        p = Path(self.file_path).resolve()
        stat = p.stat()
        return {
            "file_path": str(p),
            "name": p.name,
            "extension": p.suffix.lower(),
            "size_bytes": stat.st_size,
            "modified": self.last_modified.isoformat(),
        }

    def update_last_modified(self):
        self.last_modified = datetime.now()

    def get_file_info(self) -> Dict:
        return self._serialize_file_info()

    def get_pptx_object(self) -> Presentation:
        return self.pptx_object

    def save(self, file_path: str = None):
        if file_path:
            self.pptx_object.save(file_path)

        self.pptx_object.save(self.file_path)
        self.update_last_modified()

    def get_slides(self):
        return self.pptx_object.slides