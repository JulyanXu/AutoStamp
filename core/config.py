import json
import os
from dataclasses import dataclass, field, asdict


@dataclass
class AppConfig:
    stamp_path: str = ""
    stamp_scale: int = 100
    stamp_opacity: int = 100
    stamp_x: float = 50.0
    stamp_y: float = 50.0
    page_mode: str = "first"  # "first", "last", "all", "custom"
    custom_pages: str = ""
    output_dir: str = ""

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "AppConfig":
        if not os.path.exists(path):
            return cls()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def get_page_indices(self, total_pages: int) -> list[int]:
        if self.page_mode == "first":
            return [0]
        if self.page_mode == "last":
            return [total_pages - 1]
        if self.page_mode == "all":
            return list(range(total_pages))
        # custom: parse "1,3-5,7"
        indices = []
        for part in self.custom_pages.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start, end = part.split("-", 1)
                for i in range(int(start), int(end) + 1):
                    idx = i - 1  # 1-based to 0-based
                    if 0 <= idx < total_pages:
                        indices.append(idx)
            else:
                idx = int(part) - 1
                if 0 <= idx < total_pages:
                    indices.append(idx)
        return sorted(set(indices))
