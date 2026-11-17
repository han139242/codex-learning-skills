"""Compare every rendered PNG page by exact pixels; does not judge visual quality."""
import argparse
import json
import re
from pathlib import Path
from PIL import Image, ImageChops

def inventory(directory, pattern, expected_pages=None):
    directory = Path(directory).resolve(strict=True)
    result = {}
    for path in directory.glob(pattern):
        match = re.search(r"(\d+)$", path.stem)
        if match is None:
            raise ValueError("Page filename must end in its number: " + path.name)
        number = int(match.group(1))
        if number in result:
            raise ValueError("Duplicate page number: " + str(number))
        result[number] = path
    if not result:
        raise ValueError("No PNG pages matched in " + str(directory))
    if sorted(result) != list(range(1, max(result) + 1)):
        raise ValueError("Page sequence is incomplete in " + str(directory))
    if expected_pages is not None:
        if expected_pages < 1:
            raise ValueError("Expected page count must be positive")
        if sorted(result) != list(range(1, expected_pages + 1)):
            raise ValueError(f"PNG inventory does not match PDF page count: expected {expected_pages}, found {len(result)} in {directory}")
    return result

def compare(before, after):
    result = {"before_pages": len(before), "after_pages": len(after), "changed": [],
              "added": sorted(set(after) - set(before)), "removed": sorted(set(before) - set(after)), "unchanged": []}
    for number in sorted(set(before) & set(after)):
        with Image.open(before[number]) as image_a, Image.open(after[number]) as image_b:
            a, b = image_a.convert("RGBA"), image_b.convert("RGBA")
            if a.size != b.size:
                result["changed"].append({"page": number, "reason": "dimensions", "before": a.size, "after": b.size})
                continue
            # Compare RGBA channels independently: getbbox on RGBA alone can miss RGB changes when alpha is unchanged.
            boxes = [channel.getbbox() for channel in ImageChops.difference(a, b).split()]
            boxes = [box for box in boxes if box is not None]
            if boxes:
                box = [min(p[0] for p in boxes), min(p[1] for p in boxes), max(p[2] for p in boxes), max(p[3] for p in boxes)]
                result["changed"].append({"page": number, "reason": "pixels", "box": box})
            else:
                result["unchanged"].append(number)
    result["needs_review"] = sorted(set(result["added"]) | {p["page"] for p in result["changed"]})
    result["visual_quality_judged"] = False
    return result

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("--before-glob", default="page-*.png")
    parser.add_argument("--after-glob", default="page-*.png")
    parser.add_argument("--before-pages", type=int, help="Actual page count of the before PDF")
    parser.add_argument("--after-pages", type=int, help="Actual page count of the after PDF")
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    if args.out.exists():
        parser.error("Output already exists; choose a new temporary filename")
    result = compare(inventory(args.before, args.before_glob, args.before_pages),
                     inventory(args.after, args.after_glob, args.after_pages))
    with args.out.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2)
    print(json.dumps(result, ensure_ascii=False))

if __name__ == "__main__":
    main()
