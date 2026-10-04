"""
PDF to Markdown Converter
Batch-converts text-based PDFs to Markdown using deterministic layout
heuristics (font sizes, styles, positions). No AI and no OCR - scanned
pages are detected, reported as warnings, and skipped.

Not supported: scanned PDFs (needs OCR), multi-column layouts, image
extraction (images become placeholder comments).
"""

import re
from collections import Counter
from pathlib import Path

P2M_BULLETS = "\u2022\u25aa\u25e6\u25cf\u2023\u00b7\u2219"
P2M_BULLET_WORD_RE = re.compile(rf"^[{P2M_BULLETS}]")
P2M_CID_WORD_RE = re.compile(r"^\(cid:\d+\)$")
P2M_ORDERED_RE = re.compile(r"^(\d{1,3})[.)](?!\d)")
P2M_STRIP_BULLET_RE = re.compile(rf"^(?:[{P2M_BULLETS}]|-|\(cid:\d+\))\s*")
P2M_STRIP_ORDERED_RE = re.compile(r"^\d{1,3}[.)]\s*")
P2M_LIGATURES = {
    "\ufb00": "ff",
    "\ufb01": "fi",
    "\ufb02": "fl",
    "\ufb03": "ffi",
    "\ufb04": "ffl",
}
P2M_SENTENCE_END = (".", "!", "?", ":", ";")


def p2m_normalize(text: str) -> str:
    for lig, repl in P2M_LIGATURES.items():
        text = text.replace(lig, repl)
    return text


def p2m_font_flags(fontname: str):
    name = fontname.split("+")[-1].lower()
    bold = "bold" in name
    italic = "italic" in name or "oblique" in name
    mono = "courier" in name or "mono" in name or "consol" in name
    return bold, italic, mono


def p2m_page_links(page):
    links = []
    for annot in getattr(page, "hyperlinks", None) or []:
        uri = annot.get("uri")
        try:
            rect = (annot["x0"], annot["top"], annot["x1"], annot["bottom"])
        except (KeyError, TypeError):
            continue
        if uri and all(v is not None for v in rect):
            links.append((rect, uri))
    return links


def p2m_page_words(page):
    words = page.extract_words(extra_attrs=["fontname", "size"])
    links = p2m_page_links(page)
    out = []
    for w in words:
        if w.get("upright") is False:
            continue
        text = p2m_normalize(w.get("text", ""))
        if not text.strip():
            continue
        bold, italic, mono = p2m_font_flags(w.get("fontname") or "")
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        uri = None
        for (x0, top, x1, bottom), u in links:
            if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bottom + 1:
                uri = u
                break
        out.append({
            "text": text,
            "x0": w["x0"],
            "x1": w["x1"],
            "top": w["top"],
            "bottom": w["bottom"],
            "size": w.get("size") or 10.0,
            "bold": bold,
            "italic": italic,
            "mono": mono,
            "uri": uri,
        })
    return out


def p2m_word_in_tables(word, tables) -> bool:
    cx = (word["x0"] + word["x1"]) / 2
    cy = (word["top"] + word["bottom"]) / 2
    for t in tables:
        x0, top, x1, bottom = t["bbox"]
        if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bottom + 1:
            return True
    return False


def p2m_group_lines(words):
    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines = []
    for w in words:
        target = None
        for line in lines[-4:]:
            overlap = min(line["bottom"], w["bottom"]) - max(line["top"], w["top"])
            min_h = min(line["bottom"] - line["top"], w["bottom"] - w["top"])
            if min_h > 0 and overlap >= 0.5 * min_h:
                target = line
                break
        if target is None:
            target = {"words": [], "top": w["top"], "bottom": w["bottom"],
                      "x0": w["x0"], "x1": w["x1"]}
            lines.append(target)
        target["words"].append(w)
        target["top"] = min(target["top"], w["top"])
        target["bottom"] = max(target["bottom"], w["bottom"])
        target["x0"] = min(target["x0"], w["x0"])
        target["x1"] = max(target["x1"], w["x1"])
    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
    lines.sort(key=lambda l: (l["top"], l["x0"]))
    return lines


def p2m_line_stats(line):
    sizes = Counter()
    bold_chars = mono_chars = total = 0
    for w in line["words"]:
        n = len(w["text"])
        sizes[round(w["size"] * 2) / 2] += n
        total += n
        if w["bold"]:
            bold_chars += n
        if w["mono"]:
            mono_chars += n
    dom_size = sizes.most_common(1)[0][0] if sizes else 0.0
    bold_frac = bold_chars / total if total else 0.0
    mono_frac = mono_chars / total if total else 0.0
    return dom_size, bold_frac, mono_frac


def p2m_line_markdown(line) -> str:
    segments = []
    prev_style = None
    for w in line["words"]:
        style = (w["bold"], w["italic"], w["mono"], w["uri"])
        if segments and style == prev_style:
            segments[-1]["words"].append(w["text"])
        else:
            segments.append({"style": style, "words": [w["text"]]})
        prev_style = style
    parts = []
    for i, seg in enumerate(segments):
        if i:
            parts.append(" ")
        bold, italic, mono, uri = seg["style"]
        text = " ".join(seg["words"]).strip()
        if not text:
            continue
        if uri:
            parts.append(f"[{text}]({uri})")
        elif mono:
            parts.append(f"`{text}`")
        elif bold and italic:
            parts.append(f"***{text}***")
        elif bold:
            parts.append(f"**{text}**")
        elif italic:
            parts.append(f"*{text}*")
        else:
            parts.append(text)
    return "".join(parts)


def p2m_code_line_text(line) -> str:
    parts = []
    prev = None
    for w in line["words"]:
        if prev is not None:
            gap = w["x0"] - prev["x1"]
            if gap < 0.2 * w["size"]:
                spaces = 0
            else:
                spaces = max(1, int(round(gap / (0.6 * w["size"]))))
            parts.append(" " * spaces)
        parts.append(w["text"])
        prev = w
    return "".join(parts)


def p2m_list_kind(line):
    """Return (kind, marker) where kind is 'ul', 'ol', or None."""
    first = line["words"][0]["text"]
    if (first == "-" or P2M_BULLET_WORD_RE.match(first)
            or P2M_CID_WORD_RE.match(first)):
        return "ul", "-"
    m = P2M_ORDERED_RE.match(first)
    if m:
        return "ol", f"{m.group(1)}."
    return None, None


def p2m_cluster_levels(values, tolerance=8.0):
    starts = []
    for v in sorted(set(values)):
        if not starts or v - starts[-1] > tolerance:
            starts.append(v)

    def level_of(x):
        return sum(1 for s in starts if s < x - 4)
    return level_of


def p2m_heading_level(para_lines, body_size):
    first = para_lines[0]
    dom_size, bold_frac, _ = p2m_line_stats(first)
    ratio = dom_size / body_size if body_size else 1.0
    text_len = sum(len(w["text"]) for line in para_lines for w in line["words"])
    if len(para_lines) <= 2 and ratio >= 1.7:
        return 1
    if len(para_lines) <= 2 and ratio >= 1.4:
        return 2
    if len(para_lines) <= 2 and ratio >= 1.18:
        return 3
    if (len(para_lines) == 1 and bold_frac >= 0.9 and text_len <= 75
            and 0.95 <= ratio < 1.18):
        return 4
    return 0


def p2m_join_paragraph(para_lines) -> str:
    pieces = [p2m_line_markdown(line) for line in para_lines]
    text = pieces[0]
    for piece in pieces[1:]:
        if text.endswith("-") and piece[:1].islower():
            text = text[:-1] + piece
        else:
            text += " " + piece
    return re.sub(r"\s+", " ", text).strip()


def p2m_same_paragraph(prev_line, next_line, body_size, doc_right) -> bool:
    _, _, prev_mono = p2m_line_stats(prev_line)
    if prev_mono >= 0.6:
        return False
    gap = next_line["top"] - prev_line["bottom"]
    if gap > 0.75 * max(body_size, 8.0):
        return False
    if next_line["x0"] - prev_line["x0"] >= 0.9 * body_size:
        return False
    if prev_line["x1"] <= doc_right - 1.8 * body_size:
        if not p2m_line_markdown(prev_line).rstrip().endswith("-"):
            return False
    return True


def p2m_lines_to_blocks(lines, body_size, doc_right):
    level_of = p2m_cluster_levels(
        line["words"][0]["x0"] for line in lines
        if p2m_list_kind(line)[0]
    )
    blocks = []

    def line_kind(line):
        if p2m_line_stats(line)[2] >= 0.6:
            return "code"
        if p2m_list_kind(line)[0]:
            return "list"
        return "plain"

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        kind = line_kind(line)

        if kind == "list":
            ordered, marker = p2m_list_kind(line)
            level = level_of(line["words"][0]["x0"])
            content_x0 = line["words"][1]["x0"] if len(line["words"]) > 1 else line["x0"]
            item_lines = [line]
            j = i + 1
            while j < n and line_kind(lines[j]) == "plain":
                gap = lines[j]["top"] - lines[j - 1]["bottom"]
                if lines[j]["x0"] < content_x0 - 4 or gap > 0.75 * max(body_size, 8.0):
                    break
                item_lines.append(lines[j])
                j += 1
            text = p2m_join_paragraph(item_lines)
            strip_re = P2M_STRIP_ORDERED_RE if ordered == "ol" else P2M_STRIP_BULLET_RE
            text = strip_re.sub("", text, count=1).strip()
            if text:
                blocks.append({"kind": "li", "ordered": ordered == "ol",
                               "level": level, "marker": marker,
                               "text": text, "top": line["top"]})
            i = j

        elif kind == "code":
            code_lines = [line]
            j = i + 1
            while j < n and line_kind(lines[j]) == "code":
                gap = lines[j]["top"] - lines[j - 1]["bottom"]
                if gap > 1.5 * max(body_size, 8.0):
                    break
                code_lines.append(lines[j])
                j += 1
            base_x0 = min(cl["x0"] for cl in code_lines)
            texts = []
            for cl in code_lines:
                indent = int(round((cl["x0"] - base_x0) / (0.6 * max(body_size, 8.0))))
                texts.append(" " * indent + p2m_code_line_text(cl))
            blocks.append({"kind": "code", "text": "\n".join(texts),
                           "top": line["top"]})
            i = j

        else:
            para_lines = [line]
            j = i + 1
            while j < n and line_kind(lines[j]) == "plain":
                if not p2m_same_paragraph(lines[j - 1], lines[j], body_size, doc_right):
                    break
                para_lines.append(lines[j])
                j += 1
            text = p2m_join_paragraph(para_lines)
            if text:
                level = p2m_heading_level(para_lines, body_size)
                if level:
                    _, bold_frac, _ = p2m_line_stats(para_lines[0])
                    if bold_frac >= 0.9:
                        text = text.replace("**", "")
                    blocks.append({"kind": "h", "level": level,
                                   "text": text, "top": line["top"]})
                else:
                    blocks.append({"kind": "p", "text": text, "top": line["top"]})
            i = j

    return blocks


def p2m_render_table(rows) -> str:
    if not rows:
        return ""
    width = max(len(r) for r in rows)

    def cell(value):
        if value is None:
            return ""
        return (str(value)
                .replace("\r", "")
                .replace("\n", "<br>")
                .replace("|", "\\|")
                .strip())

    grid = [[cell(r[k] if k < len(r) else None) for k in range(width)] for r in rows]
    lines = [
        "| " + " | ".join(grid[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    lines += ["| " + " | ".join(r) + " |" for r in grid[1:]]
    return "\n".join(lines)


def p2m_zone_signature(words) -> str:
    text = " ".join(w["text"] for w in sorted(words, key=lambda w: w["x0"]))
    return re.sub(r"\d+", "#", text).strip().lower()


def p2m_strip_running_heads(pages_words, page_meta):
    n = len(pages_words)
    if n < 3:
        return pages_words
    zones = []
    hits = Counter()
    for words, meta in zip(pages_words, page_meta):
        h = meta["height"]
        limit = min(90.0, h * 0.12)
        top = [w for w in words if w["top"] < limit]
        bottom = [w for w in words if w["bottom"] > h - limit]
        zones.append((top, bottom))
        for zone in (top, bottom):
            if 0 < len(zone) <= 25:
                hits[p2m_zone_signature(zone)] += 1
    kept = []
    for i, (words, meta) in enumerate(zip(pages_words, page_meta)):
        h = meta["height"]
        limit = min(90.0, h * 0.12)
        top, bottom = zones[i]
        drop = set()
        for zone in (top, bottom):
            if 0 < len(zone) <= 25 and hits[p2m_zone_signature(zone)] >= 0.6 * n:
                drop.update(id(w) for w in zone)
        kept.append([w for w in words if id(w) not in drop])
    return kept


def p2m_render(doc_blocks) -> str:
    out = []
    list_buf = None

    def flush_list():
        nonlocal list_buf
        if list_buf:
            lines = []
            for level, marker, text in list_buf["items"]:
                lines.append("    " * level + marker + " " + text)
            out.append(("list", "\n".join(lines)))
            list_buf = None

    for block in doc_blocks:
        kind = block["kind"]
        if kind == "li":
            if list_buf is None or list_buf["ordered"] != block["ordered"]:
                flush_list()
                list_buf = {"ordered": block["ordered"], "items": []}
            list_buf["items"].append((block["level"], block["marker"], block["text"]))
        else:
            flush_list()
            if kind == "h":
                out.append(("h", "#" * block["level"] + " " + block["text"]))
            elif kind == "code":
                out.append(("code", "```\n" + block["text"] + "\n```"))
            else:
                out.append((kind, block["text"]))
    flush_list()

    merged = []
    for kind, text in out:
        if (merged and merged[-1][0] == "p" and kind == "p"
                and not merged[-1][1].rstrip().endswith(P2M_SENTENCE_END)
                and text.lstrip()[:1].islower()):
            merged[-1] = ("p", merged[-1][1].rstrip() + " " + text.lstrip())
        else:
            merged.append((kind, text))
    body = "\n\n".join(text for _, text in merged).strip()
    return body + "\n" if body else ""


def p2m_convert(pdf_path):
    """Convert one PDF. Returns (markdown, warnings). Raises ValueError when
    the file has no usable text layer (scanned/encrypted)."""
    import pdfplumber

    warnings = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        pages_words = []
        page_meta = []
        for pno, page in enumerate(pdf.pages, 1):
            rotation = getattr(page, "rotation", 0) or 0
            if rotation:
                warnings.append(f"page {pno} is rotated {rotation} degrees; "
                                f"layout may be imperfect")
            words = p2m_page_words(page)
            tables = []
            try:
                for table in page.find_tables():
                    tables.append({"bbox": table.bbox, "rows": table.extract()})
            except Exception as exc:
                warnings.append(f"page {pno}: table detection failed ({exc})")
            if tables:
                words = [w for w in words if not p2m_word_in_tables(w, tables)]
            chars = sum(len(w["text"]) for w in words)
            scanned = chars < 20 and len(page.images) > 0
            if scanned:
                warnings.append(f"page {pno} looks scanned (no text layer); "
                                f"skipped - OCR is not supported")
            pages_words.append(words)
            page_meta.append({"height": page.height, "tables": tables,
                              "images": page.images, "scanned": scanned})

    if not pages_words:
        raise ValueError("PDF has no pages")
    if all(meta["scanned"] or not words
           for words, meta in zip(pages_words, page_meta)):
        raise ValueError("PDF appears to be scanned (no extractable text). "
                         "OCR is not supported.")

    pages_words = p2m_strip_running_heads(pages_words, page_meta)

    size_counter = Counter()
    x1_values = []
    for words in pages_words:
        for w in words:
            size_counter[round(w["size"] * 2) / 2] += len(w["text"])
            x1_values.append(w["x1"])
    body_size = size_counter.most_common(1)[0][0] if size_counter else 10.0
    doc_right = sorted(x1_values)[int(0.95 * len(x1_values))] if x1_values else 612.0

    doc_blocks = []
    for pno, (words, meta) in enumerate(zip(pages_words, page_meta), 1):
        entries = []
        if meta["scanned"]:
            entries.append((0, 0, {"kind": "raw",
                                   "text": f"<!-- page {pno} skipped: scanned "
                                           f"(no text layer) -->"}))
        for img in meta["images"]:
            w = img["x1"] - img["x0"]
            h = img["bottom"] - img["top"]
            entries.append((img["top"], 1, {"kind": "raw",
                           "text": f"<!-- image: page {pno}, {w:.0f}x{h:.0f}pt -->"}))
        for table in meta["tables"]:
            md = p2m_render_table(table["rows"])
            if md:
                entries.append((table["bbox"][1], 2, {"kind": "table", "text": md}))
        if words:
            lines = p2m_group_lines(words)
            for block in p2m_lines_to_blocks(lines, body_size, doc_right):
                entries.append((block["top"], 3, block))
        entries.sort(key=lambda e: (e[0], e[1]))
        doc_blocks.extend(entry[2] for entry in entries)

    return p2m_render(doc_blocks), warnings


def convert_all_in_folder(folder: Path, log=print) -> None:
    """Convert every PDF in a folder to a sibling .md file."""
    pdfs = [
        p for p in sorted(folder.iterdir())
        if p.is_file() and p.suffix.lower() == ".pdf" and not p.name.startswith("~$")
    ]
    if not pdfs:
        log("No PDF files found.")
        return

    converted = skipped = failed = 0
    failures = []
    for pdf_path in pdfs:
        out_path = pdf_path.with_suffix(".md")
        if out_path.exists():
            log(f"Skipping (already exists): {pdf_path.name}")
            skipped += 1
            continue
        log(f"Converting: {pdf_path.name}")
        try:
            markdown, warnings = p2m_convert(pdf_path)
            for warning in warnings:
                log(f"  warning: {warning}")
            if markdown.strip():
                out_path.write_text(markdown, encoding="utf-8")
                log(f"  saved: {out_path.name}")
                converted += 1
            else:
                log("  FAILED: no extractable text")
                failed += 1
                failures.append((pdf_path.name, "no extractable text"))
        except ValueError as exc:
            log(f"  FAILED: {exc}")
            failed += 1
            failures.append((pdf_path.name, str(exc)))
        except Exception as exc:
            log(f"  FAILED: {exc}")
            failed += 1
            failures.append((pdf_path.name, str(exc)))

    log("\n— Summary —")
    log(f"Converted: {converted}")
    if skipped:
        log(f"Skipped (existing .md): {skipped}")
    if failures:
        log(f"Failed ({len(failures)}):")
        for name, err in failures:
            log(f"   - {name}: {err}")
    else:
        log("No failures.")


def main():
    """Run the converter on the current directory."""
    folder = Path(__file__).resolve().parent
    print(f"Folder: {folder}")
    convert_all_in_folder(folder)


if __name__ == "__main__":
    main()
