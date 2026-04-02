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
        target_path = Path(file_path) if file_path else Path(self.file_path)
        self.pptx_object.save(target_path)
        self.file_path = str(target_path)
        self.update_last_modified()

    def get_slides(self):
        return self.pptx_object.slides

class Pictogram:
    def __init__(self, name: str, image_path: str):
        self.name = name
        self.image_path = image_path

    def to_dict(self) -> Dict:
        return {
            "name": self.name
        }

class PictogramLibrary:
    def __init__(self):
        pictogram_path = Path(__file__).parent / "pictogram"
        self.pictograms: Dict[str, Pictogram] = {}
        if pictogram_path.exists() and pictogram_path.is_dir():
            for file in pictogram_path.iterdir():
                if file.is_file() and file.suffix.lower() in {".png", ".jpg", ".jpeg", ".svg"}:
                    name = file.stem
                    self.pictograms[name] = Pictogram(name, str(file))

    def get_pictogram(self, name: str) -> Pictogram:
        return self.pictograms.get(name)

    def list_pictograms(self) -> Dict[str, Pictogram]:
        return self.pictograms