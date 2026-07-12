from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def slide_number(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def text_from_xml(blob: bytes) -> list[str]:
    root = ET.fromstring(blob)
    chunks: list[str] = []
    for paragraph in root.findall(".//a:p", NS):
        runs = [t.text or "" for t in paragraph.findall(".//a:t", NS)]
        line = "".join(runs).strip()
        if line:
            chunks.append(line)
    return chunks


def extract_pptx(path: Path) -> dict:
    slides = []
    with zipfile.ZipFile(path) as zf:
        names = sorted(
            [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=slide_number,
        )
        for name in names:
            try:
                lines = text_from_xml(zf.read(name))
            except ET.ParseError:
                lines = []
            slides.append({"slide": slide_number(name), "text": lines})

        notes = []
        note_names = sorted(
            [n for n in zf.namelist() if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)],
            key=slide_number,
        )
        for name in note_names:
            try:
                lines = text_from_xml(zf.read(name))
            except ET.ParseError:
                lines = []
            if lines:
                notes.append({"slide": slide_number(name), "text": lines})
    return {"file": str(path), "slide_count": len(slides), "slides": slides, "notes": notes}


def main() -> None:
    root = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)
    pptx_files = [
        p
        for p in sorted(root.glob("*.pptx"))
        if not p.name.startswith("~$")
    ]
    data = [extract_pptx(p) for p in pptx_files]
    (out_dir / "ppt_text.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    lines: list[str] = []
    for deck in data:
        lines.append(f"# {Path(deck['file']).name} ({deck['slide_count']} slides)")
        for slide in deck["slides"]:
            if not slide["text"]:
                continue
            lines.append(f"\n## Slide {slide['slide']}")
            lines.extend(f"- {item}" for item in slide["text"])
        lines.append("")
    (out_dir / "ppt_text.txt").write_text("\n".join(lines), encoding="utf-8")
    print(f"Extracted {len(data)} decks to {out_dir}")
    for deck in data:
        nonempty = sum(1 for s in deck["slides"] if s["text"])
        print(f"{Path(deck['file']).name}: {deck['slide_count']} slides, {nonempty} with text")


if __name__ == "__main__":
    main()
