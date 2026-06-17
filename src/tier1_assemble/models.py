from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional


@dataclass(frozen=True)
class AssemblyOptions:
    overwrite: bool = False
    anchor: str = "center"
    margin: int = 0


@dataclass(frozen=True)
class AssetBinding:
    token: str
    path: Optional[Path] = None
    kind: str = "image"
    base_path: Optional[Path] = None
    overlay_path: Optional[Path] = None
    anchor: str = "bottom-right"
    margin_px: int = 50


@dataclass(frozen=True)
class AssemblyManifest:
    manifest_path: Path
    assets_dir: Path
    run_id: str
    assets: Dict[str, AssetBinding]


@dataclass
class TokenReplacement:
    token: str
    slide_number: int
    shape_name: str
    asset_path: Optional[Path] = None
    inserted: bool = False
    visible_text_token: bool = False
    warning: Optional[str] = None

    def as_dict(self) -> Dict[str, object]:
        return {
            "token": self.token,
            "slide_number": self.slide_number,
            "shape_name": self.shape_name,
            "asset_path": str(self.asset_path) if self.asset_path else None,
            "inserted": self.inserted,
            "visible_text_token": self.visible_text_token,
            "warning": self.warning,
        }


@dataclass
class AssemblyReport:
    template_path: Path
    manifest_path: Path
    output_path: Path
    run_id: str
    slides_scanned: int = 0
    tokens_found: int = 0
    assets_inserted: int = 0
    derived_assets_created: int = 0
    unresolved_tokens: List[str] = field(default_factory=list)
    unused_assets: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    replacements: List[TokenReplacement] = field(default_factory=list)

    @property
    def valid(self) -> bool:
        if self.unresolved_tokens:
            return False
        return all(replacement.inserted for replacement in self.replacements)

    def as_dict(self) -> Dict[str, object]:
        return {
            "valid": self.valid,
            "template_path": str(self.template_path),
            "manifest_path": str(self.manifest_path),
            "output_path": str(self.output_path),
            "run_id": self.run_id,
            "slides_scanned": self.slides_scanned,
            "tokens_found": self.tokens_found,
            "assets_inserted": self.assets_inserted,
            "derived_assets_created": self.derived_assets_created,
            "unresolved_tokens": list(self.unresolved_tokens),
            "unused_assets": list(self.unused_assets),
            "warnings": list(self.warnings),
            "replacements": [replacement.as_dict() for replacement in self.replacements],
        }
