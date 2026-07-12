from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def natural_key(path: str) -> tuple[int, str]:
    match = re.search(r"(\d+)", Path(path).stem)
    return (int(match.group(1)) if match else 0, path)


def extract_slide_text(zf: ZipFile, slide_name: str) -> list[str]:
    root = etree.fromstring(zf.read(slide_name))
    chunks: list[str] = []
    for shape in root.xpath(".//p:sp | .//p:graphicFrame", namespaces=NS):
        texts = []
        for t in shape.xpath(".//a:t", namespaces=NS):
            if t.text:
                clean = " ".join(t.text.split())
                if clean:
                    texts.append(clean)
        if texts:
            chunks.append(" ".join(texts))
    return chunks


def extract_pptx(path: Path) -> dict:
    with ZipFile(path) as zf:
        slides = sorted(
            [name for name in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)],
            key=natural_key,
        )
        extracted = []
        for i, slide in enumerate(slides, start=1):
            chunks = extract_slide_text(zf, slide)
            extracted.append(
                {
                    "slide": i,
                    "text": chunks,
                    "joined": "\n".join(chunks),
                }
            )
    return {"file": str(path), "slide_count": len(extracted), "slides": extracted}


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: extract_pptx_text.py <pptx_dir> <out_dir>", file=sys.stderr)
        return 2

    pptx_dir = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)

    index = []
    for pptx in sorted(pptx_dir.glob("*.pptx")):
        if pptx.name.startswith("~$"):
            continue
        data = extract_pptx(pptx)
        out_json = out_dir / f"{pptx.stem}.json"
        out_txt = out_dir / f"{pptx.stem}.txt"
        out_json.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        lines = [f"# {pptx.name}", f"幻灯片数：{data['slide_count']}", ""]
        for slide in data["slides"]:
            lines.append(f"## Slide {slide['slide']}")
            lines.extend(slide["text"])
            lines.append("")
        out_txt.write_text("\n".join(lines), encoding="utf-8")
        index.append({"name": pptx.name, "slide_count": data["slide_count"], "txt": str(out_txt)})

    (out_dir / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    for item in index:
        print(f"{item['name']}\t{item['slide_count']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
