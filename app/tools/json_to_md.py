from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable


def _escape_md(text: str) -> str:
    # Basic escaping to avoid accidental headings or formatting
    return (
        text.replace("\\", "\\\\")
        .replace("`", "\``")
        .replace("*", "\\*")
        .replace("_", "\\_")
        .replace("[", "\\[")
        .replace("]", "\\]")
        .replace("#", "\\#")
    )


def _join(values: Iterable[str]) -> str:
    return ", ".join(v for v in values if v)


def json_to_markdown(obj: dict[str, Any]) -> str:
    message = str(obj.get("message") or obj.get("Message") or "Danh sách Video")
    data = obj.get("data")
    if not isinstance(data, list):
        raise ValueError("JSON missing 'data' list")

    # Group by subject (Môn học)
    by_subject: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in data:
        if isinstance(item, dict):
            subject = str(item.get("Môn học", "Khác"))
            by_subject[subject].append(item)

    lines: list[str] = []
    lines.append(f"# {_escape_md(message)}")
    lines.append("")

    for subject in sorted(by_subject.keys(), key=lambda s: s.lower()):
        lines.append(f"## Môn học: {_escape_md(subject)}")
        lines.append("")

        for item in by_subject[subject]:
            title = str(item.get("Tiêu đề", "(Không tiêu đề)"))
            code = str(item.get("Mã", ""))
            desc = str(item.get("Mô tả", ""))
            hashtags = item.get("Hashtag") or []
            if isinstance(hashtags, list):
                hashtag_str = _join([str(x) for x in hashtags])
            else:
                hashtag_str = str(hashtags)
            id_course = str(item.get("id_course", ""))

            fields = item.get("Lĩnh vực(Optional)") or item.get("Lĩnh vực") or []
            vn_names: list[str] = []
            en_names: list[str] = []
            if isinstance(fields, list):
                for f in fields:
                    if isinstance(f, dict):
                        vn = f.get("FIELD_OF_STUDY_NAME")
                        en = f.get("FIELD_OF_STUDY_NAME_EN")
                        if vn:
                            vn_names.append(str(vn))
                        if en:
                            en_names.append(str(en))

            # Top bullet: title only
            lines.append(f"- {_escape_md(title)}")

            # Sub bullets (ordered): ID -> Mô tả -> Hashtag -> id_course -> Lĩnh vực -> Lĩnh vực (EN)
            if code:
                lines.append(f"  - ID: {_escape_md(code)}")
            if desc:
                lines.append(f"  - Mô tả: {_escape_md(desc)}")
            if hashtags:
                lines.append(f"  - Hashtag: {_escape_md(hashtag_str)}")
            if id_course:
                lines.append(f"  - id_course: `{_escape_md(id_course)}`")
            if vn_names:
                lines.append(f"  - Lĩnh vực: {_escape_md(_join(vn_names))}")
            if en_names:
                lines.append(f"  - Lĩnh vực (EN): {_escape_md(_join(en_names))}")

        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Convert JSON (videos) to Markdown list grouped by subject")
    parser.add_argument("--input", "-i", required=True, help="Path to input JSON file")
    parser.add_argument("--output", "-o", required=True, help="Path to output Markdown file")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        raise SystemExit(f"Input file not found: {in_path}")

    raw = in_path.read_text(encoding="utf-8-sig")
    obj = json.loads(raw)
    if not isinstance(obj, dict):
        raise SystemExit("Top-level JSON must be an object")

    md = json_to_markdown(obj)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(md, encoding="utf-8")
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
