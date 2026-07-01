"""
text_editor.py
==============
Read and replace text content in .pptx template files via slide XML.

Quick usage:
    editor = PptxTextEditor("Template.pptx")
    editor.list_texts()                                     # see all text shapes
    editor.replace_by_shape(1, "TextBox 10", "New text")   # by slide + shape name
    editor.replace_placeholder("{{brand}}", "Masan")       # find-and-replace across all slides
    editor.replace_all({"{{brand}}": "Masan", "{{date}}": "01/06/2025"})
    editor.save("Template_updated.pptx")
"""

import os
import re
import shutil
import zipfile
import tempfile
from copy import deepcopy
from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_REL_TYPE = f"{R}/hyperlink"


class PptxTextEditor:
    def __init__(self, pptx_path: str):
        self.pptx_path = pptx_path
        self.tmp_dir = tempfile.mkdtemp(prefix="pptx_text_")
        shutil.unpack_archive(pptx_path, self.tmp_dir, "zip")
        self._slide_trees: dict[int, etree._ElementTree] = {}
        self._slides_dir = os.path.join(self.tmp_dir, "ppt", "slides")
        print(f"✅ Loaded: {pptx_path}")
        print(f"   Slides found: {len(self._get_slide_nums())}")


    def list_texts(self, slide_num: int = None) -> list[dict]:
        """
        List all text shapes in the presentation (or a single slide).

        Returns a list of dicts: {slide, shape_name, shape_id, text}
        """
        results = []
        nums = [slide_num] if slide_num else self._get_slide_nums()
        for n in nums:
            root = self._get_tree(n).getroot()
            for sp in root.iter(f"{{{P}}}sp"):
                name, sid = self._shape_attrs(sp)
                text = self._collect_text(sp)
                if text:
                    results.append({"slide": n, "shape_name": name, "shape_id": sid, "text": text[:120]})
                    print(f"  slide{n}  [{sid}] '{name}' → {text[:80]!r}")
        return results

    def replace_by_shape(self, slide_num: int, shape_name: str, new_text: str,
                         bold_first_line: bool = False) -> bool:
        """
        Replace all text in a named shape on a specific slide.

        Preserves the formatting of the first run; clears remaining runs.
        Extra \\n lines beyond template paragraphs are added as new paragraphs
        cloned from the first template paragraph.
        If bold_first_line=True the first paragraph is rendered bold and the
        rest normal-weight.
        Returns True if the shape was found and updated.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            name, _ = self._shape_attrs(sp)
            if name == shape_name:
                self._set_shape_text(sp, new_text, bold_first_line)
                print(f"✅ slide{slide_num} '{shape_name}' → {new_text[:60]!r}")
                return True
        print(f"⚠️  Shape '{shape_name}' not found on slide {slide_num}")
        return False

    def replace_placeholder(self, old_text: str, new_text: str, slide_num: int = None) -> int:
        """
        Find-and-replace a placeholder string across all slides (or one slide).

        Handles the case where a placeholder is split across adjacent runs within
        the same paragraph by first joining paragraph text, replacing, then
        writing the result back to the first run.

        Returns the number of replacements made.
        """
        count = 0
        nums = [slide_num] if slide_num else self._get_slide_nums()
        for n in nums:
            root = self._get_tree(n).getroot()
            for para in root.iter(f"{{{A}}}p"):
                count += self._replace_in_para(para, old_text, new_text)
        if count:
            print(f"✅ Replaced {count}x  {old_text!r} → {new_text!r}")
        return count

    def replace_all(self, replacements: dict, slide_num: int = None) -> dict:
        """
        Apply multiple find-and-replace pairs in one pass.

        Args:
            replacements: {old_text: new_text, ...}
            slide_num:    Apply only to this slide; None = all slides.

        Returns:
            {old_text: count_of_replacements}
        """
        counts = {k: 0 for k in replacements}
        nums = [slide_num] if slide_num else self._get_slide_nums()
        for n in nums:
            root = self._get_tree(n).getroot()
            for para in root.iter(f"{{{A}}}p"):
                for old, new in replacements.items():
                    counts[old] += self._replace_in_para(para, old, new)
        for old, cnt in counts.items():
            if cnt:
                print(f"✅ Replaced {cnt}x  {old!r} → {replacements[old]!r}")
        return counts

    def fill_table_on_slide(self, slide_num: int, posts: list[dict]) -> bool:
        """Populate the first table found on a slide with top-posts data.

        Always clears every data row first so template data never bleeds through.
        If posts is empty, writes a placeholder message in the first data row.
        Returns True if a table was found.
        """
        root = self._get_tree(slide_num).getroot()
        tbl  = root.find(f".//{{{A}}}tbl")
        if tbl is None:
            return False
        rows      = tbl.findall(f"{{{A}}}tr")
        data_rows = rows[1:]  # skip header row

        # Always clear all data rows first
        for row in data_rows:
            for t in row.iter(f"{{{A}}}t"):
                t.text = ""

        if not posts:
            if data_rows:
                cells = data_rows[0].findall(f"{{{A}}}tc")
                if cells:
                    for t in cells[1].iter(f"{{{A}}}t"):
                        t.text = "Không có bài đăng nổi bật được ghi nhận"
                        break
            print(f"✅ Cleared table on slide {slide_num} (no posts)")
            return True

        for row_idx, row in enumerate(data_rows):
            if row_idx >= len(posts):
                break
            post   = posts[row_idx]
            cells  = row.findall(f"{{{A}}}tc")
            values = [
                str(post.get("Rank", "")),
                " ".join((post.get("Title") or post.get("Content") or post.get("Description", "")).split()[:20]),
                str(post.get("Interaction", 0)),
                post.get("SiteName", ""),
                "100%" if post.get("Sentiment") == "Positive" else "0%",
                "100%" if post.get("Sentiment") == "Neutral" else "0%",
                "100%" if post.get("Sentiment") == "Negative" else "0%",
            ]
            for ci, cell in enumerate(cells):
                if ci >= len(values):
                    break
                for t in cell.iter(f"{{{A}}}t"):
                    t.text = values[ci]
                    break
        print(f"✅ Filled table on slide {slide_num} with {min(len(posts), len(data_rows))} posts")
        return True

    def replace_run_in_shape(
        self, slide_num: int, shape_name: str,
        para_idx: int, run_idx: int, new_text: str,
    ) -> bool:
        """Replace text in a specific run (para_idx, run_idx) of a named shape.

        Preserves all run formatting (font size, bold, color, etc.).
        Returns True if the run was found and updated.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            name, _ = self._shape_attrs(sp)
            if name != shape_name:
                continue
            txBody = sp.find(f"{{{P}}}txBody")
            if txBody is None:
                return False
            paras = txBody.findall(f"{{{A}}}p")
            if para_idx >= len(paras):
                return False
            runs = paras[para_idx].findall(f"{{{A}}}r")
            if run_idx >= len(runs):
                return False
            t = runs[run_idx].find(f"{{{A}}}t")
            if t is not None:
                t.text = new_text
                print(f"✅ slide{slide_num} '{shape_name}'[p{para_idx}r{run_idx}] → {new_text!r}")
                return True
        print(f"⚠️  Shape '{shape_name}' not found on slide {slide_num}")
        return False

    def get_chart_rels(self, slide_num: int) -> dict[str, str]:
        """Return {rId: chart_filename} for all chart relationships on this slide."""
        rels_path = os.path.join(self._slides_dir, "_rels", f"slide{slide_num}.xml.rels")
        if not os.path.exists(rels_path):
            return {}
        try:
            rels_root = etree.parse(rels_path).getroot()
        except Exception:
            return {}
        result = {}
        for rel in rels_root:
            if "chart" in rel.get("Type", ""):
                fname = rel.get("Target", "").split("/")[-1]
                result[rel.get("Id", "")] = fname
        return result

    def set_chart_frame_transform(
        self, slide_num: int, chart_filename: str,
        x: int | None = None, y: int | None = None,
        cx: int | None = None, cy: int | None = None,
    ) -> bool:
        """Resize/reposition the graphic frame that contains the given chart file.
        Pass None for any dimension to leave it unchanged.
        Returns True if the frame was found and updated.
        """
        rels = self.get_chart_rels(slide_num)
        rid = {v: k for k, v in rels.items()}.get(chart_filename)
        if not rid:
            return False
        root = self._get_tree(slide_num).getroot()
        for gf in root.iter(f"{{{P}}}graphicFrame"):
            chart_el = gf.find(f".//{{{C}}}chart")
            if chart_el is None:
                continue
            if chart_el.get(f"{{{R}}}id") != rid:
                continue
            xfrm = gf.find(f"{{{P}}}xfrm")
            if xfrm is None:
                return False
            off = xfrm.find(f"{{{A}}}off")
            ext = xfrm.find(f"{{{A}}}ext")
            if off is not None:
                if x  is not None: off.set("x",  str(x))
                if y  is not None: off.set("y",  str(y))
            if ext is not None:
                if cx is not None: ext.set("cx", str(cx))
                if cy is not None: ext.set("cy", str(cy))
            return True
        return False

    def set_shape_transform_by_name(
        self, slide_num: int, shape_name: str,
        x: int | None = None, y: int | None = None,
        cx: int | None = None, cy: int | None = None,
    ) -> bool:
        """Resize/reposition a text shape by its name.
        Pass None for any dimension to leave it unchanged.
        Returns True if the shape was found and updated.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            cNvPr = sp.find(f".//{{{P}}}cNvPr")
            if cNvPr is None or cNvPr.get("name") != shape_name:
                continue
            xfrm = sp.find(f".//{{{A}}}xfrm")
            if xfrm is None:
                return False
            off = xfrm.find(f"{{{A}}}off")
            ext = xfrm.find(f"{{{A}}}ext")
            if off is not None:
                if x  is not None: off.set("x",  str(x))
                if y  is not None: off.set("y",  str(y))
            if ext is not None:
                if cx is not None: ext.set("cx", str(cx))
                if cy is not None: ext.set("cy", str(cy))
            return True
        return False

    def append_hyperlink_run_to_shape(
        self,
        slide_num: int,
        shape_name: str,
        url_per_runnable_idx: dict,
        link_text: str = " (URL)",
    ) -> bool:
        """For each targeted runnable paragraph in a named shape:
        1. Strip any existing <a:hlinkClick> from all runs (clears template links).
        2. Append a new run containing link_text with a hyperlink to the given URL.

        Only runs with non-empty, non-null URLs are processed.
        Returns True if the shape was found.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            name, _ = self._shape_attrs(sp)
            if name != shape_name:
                continue
            txBody = sp.find(f"{{{P}}}txBody")
            if txBody is None:
                return False
            runnable_idx = 0
            for para in txBody.findall(f"{{{A}}}p"):
                runs = para.findall(f"{{{A}}}r")
                if not runs:
                    continue
                url = url_per_runnable_idx.get(runnable_idx, "")
                if url and url not in ("None", "nan"):
                    # Strip old hyperlinks from every existing run
                    for run in runs:
                        rPr = run.find(f"{{{A}}}rPr")
                        if rPr is not None:
                            for hlink in rPr.findall(f"{{{A}}}hlinkClick"):
                                rPr.remove(hlink)
                    # Build new hyperlinked run (clone first run's rPr for style)
                    rid = self._add_slide_hyperlink_rel(slide_num, url)
                    first_rPr = runs[0].find(f"{{{A}}}rPr")
                    new_rPr = deepcopy(first_rPr) if first_rPr is not None else etree.Element(f"{{{A}}}rPr")
                    for hlink in new_rPr.findall(f"{{{A}}}hlinkClick"):
                        new_rPr.remove(hlink)
                    hlink_el = etree.SubElement(new_rPr, f"{{{A}}}hlinkClick")
                    hlink_el.set(f"{{{R}}}id", rid)
                    new_run = etree.SubElement(para, f"{{{A}}}r")
                    new_run.append(new_rPr)
                    new_t = etree.SubElement(new_run, f"{{{A}}}t")
                    new_t.text = link_text
                runnable_idx += 1
            return True
        print(f"⚠️  Shape '{shape_name}' not found on slide {slide_num}")
        return False

    def hyperlink_para_content(
        self,
        slide_num: int,
        shape_name: str,
        url_per_runnable_idx: dict,
        prefix_end: str = ": ",
    ) -> bool:
        """Hyperlink only the content part of "Prefix: content" paragraphs.

        For each targeted runnable paragraph:
        1. Strip existing <a:hlinkClick> from all runs.
        2. Split the first run at the first occurrence of prefix_end.
        3. Keep the prefix text in the original run (no hyperlink).
        4. Insert a new run after it containing the content text with a hyperlink.
        If prefix_end is not found, the entire run text is treated as content.
        Returns True if the shape was found.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            name, _ = self._shape_attrs(sp)
            if name != shape_name:
                continue
            txBody = sp.find(f"{{{P}}}txBody")
            if txBody is None:
                return False
            runnable_idx = 0
            for para in txBody.findall(f"{{{A}}}p"):
                runs = para.findall(f"{{{A}}}r")
                if not runs:
                    continue
                url = url_per_runnable_idx.get(runnable_idx, "")
                if url and url not in ("None", "nan"):
                    # Strip old hyperlinks
                    for run in runs:
                        rPr = run.find(f"{{{A}}}rPr")
                        if rPr is not None:
                            for hlink in rPr.findall(f"{{{A}}}hlinkClick"):
                                rPr.remove(hlink)
                    # Split first run at first prefix_end
                    first_run = runs[0]
                    first_t = first_run.find(f"{{{A}}}t")
                    full_text = (first_t.text or "") if first_t is not None else ""
                    sep_idx = full_text.find(prefix_end)
                    if sep_idx != -1:
                        prefix_text  = full_text[:sep_idx + len(prefix_end)]
                        content_text = full_text[sep_idx + len(prefix_end):]
                    else:
                        prefix_text  = ""
                        content_text = full_text
                    if first_t is not None:
                        first_t.text = prefix_text
                    # Build content run with hyperlink
                    rid = self._add_slide_hyperlink_rel(slide_num, url)
                    first_rPr = first_run.find(f"{{{A}}}rPr")
                    new_rPr = deepcopy(first_rPr) if first_rPr is not None else etree.Element(f"{{{A}}}rPr")
                    for hlink in new_rPr.findall(f"{{{A}}}hlinkClick"):
                        new_rPr.remove(hlink)
                    hlink_el = etree.SubElement(new_rPr, f"{{{A}}}hlinkClick")
                    hlink_el.set(f"{{{R}}}id", rid)
                    new_run = etree.Element(f"{{{A}}}r")
                    new_run.append(new_rPr)
                    new_t = etree.SubElement(new_run, f"{{{A}}}t")
                    new_t.text = content_text
                    # Insert immediately after first run
                    para.insert(list(para).index(first_run) + 1, new_run)
                runnable_idx += 1
            return True
        print(f"⚠️  Shape '{shape_name}' not found on slide {slide_num}")
        return False

    def add_shape_para_hyperlinks(
        self,
        slide_num: int,
        shape_name: str,
        url_per_runnable_idx: dict,
    ) -> bool:
        """Add/replace hyperlinks on specific runnable paragraphs of a named shape.

        url_per_runnable_idx: {runnable_para_idx: url}
        Runnable paragraphs are those with at least one <a:r> (spacers skipped).
        Only entries with non-empty URLs are applied.
        Returns True if the shape was found.
        """
        root = self._get_tree(slide_num).getroot()
        for sp in root.iter(f"{{{P}}}sp"):
            name, _ = self._shape_attrs(sp)
            if name != shape_name:
                continue
            txBody = sp.find(f"{{{P}}}txBody")
            if txBody is None:
                return False
            runnable_idx = 0
            for para in txBody.findall(f"{{{A}}}p"):
                runs = para.findall(f"{{{A}}}r")
                if not runs:
                    continue
                url = url_per_runnable_idx.get(runnable_idx, "")
                if url and url not in ("None", "nan"):
                    self._add_hyperlink_to_runs(slide_num, runs, url)
                runnable_idx += 1
            return True
        print(f"⚠️  Shape '{shape_name}' not found on slide {slide_num}")
        return False

    def add_table_cell_hyperlink(
        self,
        slide_num: int,
        row_idx: int,
        col_idx: int,
        url: str,
    ) -> bool:
        """Add/replace hyperlink on all runs in a data-row table cell.

        row_idx is 0-indexed into data rows (header row skipped).
        Returns True if the cell was found.
        """
        if not url or url in ("None", "nan"):
            return False
        root = self._get_tree(slide_num).getroot()
        tbl = root.find(f".//{{{A}}}tbl")
        if tbl is None:
            return False
        rows = tbl.findall(f"{{{A}}}tr")
        data_rows = rows[1:]
        if row_idx >= len(data_rows):
            return False
        cells = data_rows[row_idx].findall(f"{{{A}}}tc")
        if col_idx >= len(cells):
            return False
        runs = list(cells[col_idx].iter(f"{{{A}}}r"))
        self._add_hyperlink_to_runs(slide_num, runs, url)
        return True

    def save(self, output_path: str = None) -> str:
        """
        Save all changes to a new .pptx file.

        Args:
            output_path: Destination path.  Default: original name + '_updated'.
        """
        if output_path is None:
            base, ext = os.path.splitext(self.pptx_path)
            output_path = f"{base}_updated{ext}"

        for slide_num, tree in self._slide_trees.items():
            path = os.path.join(self._slides_dir, f"slide{slide_num}.xml")
            tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)

        self._pack_pptx(self.tmp_dir, output_path)
        print(f"\n💾 Saved → {output_path}")
        return output_path

    def __del__(self):
        try:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        except Exception:
            pass

    # ─────────────────────────────────────────────
    # INTERNAL HELPERS
    # ─────────────────────────────────────────────

    def _get_slide_nums(self) -> list[int]:
        nums = []
        for fname in os.listdir(self._slides_dir):
            m = re.match(r"^slide(\d+)\.xml$", fname)
            if m:
                nums.append(int(m.group(1)))
        return sorted(nums)

    def _get_tree(self, slide_num: int) -> etree._ElementTree:
        if slide_num not in self._slide_trees:
            path = os.path.join(self._slides_dir, f"slide{slide_num}.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"slide{slide_num}.xml not found")
            self._slide_trees[slide_num] = etree.parse(path)
        return self._slide_trees[slide_num]

    @staticmethod
    def _shape_attrs(sp) -> tuple[str, str]:
        cNvPr = sp.find(f".//{{{P}}}cNvPr")
        if cNvPr is None:
            return "", ""
        return cNvPr.get("name", ""), cNvPr.get("id", "")

    @staticmethod
    def _collect_text(sp) -> str:
        parts = []
        for t in sp.iter(f"{{{A}}}t"):
            if t.text:
                parts.append(t.text)
        return "".join(parts).strip()

    @staticmethod
    def _set_bold(run, bold: bool):
        """Set or clear the b attribute on the run's rPr element."""
        rPr = run.find(f"{{{A}}}rPr")
        if rPr is not None:
            rPr.set("b", "1" if bold else "0")

    @staticmethod
    def _set_shape_text(sp, new_text: str, bold_first_line: bool = False):
        """Replace all text in the shape's txBody with new_text.

        Each \\n in new_text maps to a separate visible paragraph.
        Paragraphs with zero runs (spacers) are skipped so they remain
        unchanged and lines are only mapped to paragraphs that actually
        have runs.

        - If there are more lines than runnable template paragraphs, extra
          paragraphs are cloned from the last runnable paragraph (preserving
          content-area formatting, not the bold header).
        - If there are fewer lines, excess runnable paragraphs are cleared.
        - If bold_first_line=True, the first mapped paragraph gets b="1" and
          all others get b="0".
        """
        txBody = sp.find(f"{{{P}}}txBody")
        if txBody is None:
            return

        lines = new_text.split("\n")
        all_paras = txBody.findall(f"{{{A}}}p")

        if not all_paras:
            return

        # Only map to paragraphs that have at least one <a:r> (skip spacers)
        runnable = [(p_pos, para) for p_pos, para in enumerate(all_paras)
                    if para.findall(f"{{{A}}}r")]

        if not runnable:
            return

        # ── Write lines into runnable paragraphs ──────────────────────────────
        for line_idx, (_, para) in enumerate(runnable):
            runs = para.findall(f"{{{A}}}r")
            line_text = lines[line_idx] if line_idx < len(lines) else ""

            first_t = runs[0].find(f"{{{A}}}t")
            if first_t is not None:
                first_t.text = line_text
            for run in runs[1:]:
                t = run.find(f"{{{A}}}t")
                if t is not None:
                    t.text = ""

            if bold_first_line:
                is_bold = (line_idx == 0)
                for run in runs:
                    PptxTextEditor._set_bold(run, is_bold)

        # ── Add new paragraphs for lines beyond the runnable count ────────────
        if len(lines) > len(runnable):
            # Clone from the last runnable paragraph (content formatting, not header)
            _, template_para = runnable[-1]
            all_children = list(txBody)
            last_para_pos = max(
                i for i, ch in enumerate(all_children)
                if ch.tag == f"{{{A}}}p"
            )

            for i in range(len(runnable), len(lines)):
                new_para = deepcopy(template_para)
                new_runs = new_para.findall(f"{{{A}}}r")
                if new_runs:
                    first_t = new_runs[0].find(f"{{{A}}}t")
                    if first_t is not None:
                        first_t.text = lines[i]
                    for run in new_runs[1:]:
                        t = run.find(f"{{{A}}}t")
                        if t is not None:
                            t.text = ""
                    if bold_first_line:
                        for run in new_runs:
                            PptxTextEditor._set_bold(run, False)

                last_para_pos += 1
                txBody.insert(last_para_pos, new_para)

    @staticmethod
    def _replace_in_para(para, old_text: str, new_text: str) -> int:
        """Replace old_text within a paragraph, handling cross-run splits.

        Strategy:
        1. Check each run individually (fast path for simple cases).
        2. If not found per-run, join all run texts, replace, and write result
           back to the first run while clearing the rest (loses per-run formatting
           but preserves paragraph formatting — acceptable for template placeholders).

        Returns the number of replacements performed (0 or 1 per paragraph).
        """
        runs = para.findall(f"{{{A}}}r")
        if not runs:
            return 0

        # Fast path: placeholder is fully inside one run
        for run in runs:
            t = run.find(f"{{{A}}}t")
            if t is not None and t.text and old_text in t.text:
                t.text = t.text.replace(old_text, new_text)
                return 1

        # Slow path: placeholder may span multiple runs
        combined = "".join((r.find(f"{{{A}}}t").text or "") for r in runs if r.find(f"{{{A}}}t") is not None)
        if old_text not in combined:
            return 0

        replaced = combined.replace(old_text, new_text)
        first_t = runs[0].find(f"{{{A}}}t")
        if first_t is not None:
            first_t.text = replaced
        for run in runs[1:]:
            t = run.find(f"{{{A}}}t")
            if t is not None:
                t.text = ""
        return 1

    def _add_slide_hyperlink_rel(self, slide_num: int, url: str) -> str:
        """Add (or reuse) a hyperlink relationship in slide{n}.xml.rels; return rId."""
        rels_dir  = os.path.join(self._slides_dir, "_rels")
        rels_path = os.path.join(rels_dir, f"slide{slide_num}.xml.rels")

        if os.path.exists(rels_path):
            tree = etree.parse(rels_path)
            root = tree.getroot()
        else:
            os.makedirs(rels_dir, exist_ok=True)
            root = etree.Element(
                f"{{{RELS_NS}}}Relationships",
                nsmap={None: RELS_NS},
            )
            tree = etree.ElementTree(root)

        # Reuse existing relationship for the same URL
        for rel in root:
            if rel.get("Target") == url and rel.get("Type") == HYPERLINK_REL_TYPE:
                return rel.get("Id")

        # Assign next available rId
        max_id = 0
        for rel in root:
            m = re.match(r"rId(\d+)$", rel.get("Id", ""))
            if m:
                max_id = max(max_id, int(m.group(1)))

        new_rid = f"rId{max_id + 1}"
        rel_el = etree.SubElement(root, f"{{{RELS_NS}}}Relationship")
        rel_el.set("Id",         new_rid)
        rel_el.set("Type",       HYPERLINK_REL_TYPE)
        rel_el.set("Target",     url)
        rel_el.set("TargetMode", "External")

        tree.write(rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        return new_rid

    def _add_hyperlink_to_runs(self, slide_num: int, runs: list, url: str) -> None:
        """Add/replace <a:hlinkClick r:id="..."/> on every run in the list."""
        if not runs or not url:
            return
        rid = self._add_slide_hyperlink_rel(slide_num, url)
        for run in runs:
            rPr = run.find(f"{{{A}}}rPr")
            if rPr is None:
                rPr = etree.Element(f"{{{A}}}rPr")
                run.insert(0, rPr)
            for old in rPr.findall(f"{{{A}}}hlinkClick"):
                rPr.remove(old)
            hlink = etree.SubElement(rPr, f"{{{A}}}hlinkClick")
            hlink.set(f"{{{R}}}id", rid)

    @staticmethod
    def _pack_pptx(src_dir: str, output_path: str):
        tmp_zip = output_path + ".tmp.zip"
        with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(src_dir):
                dirs[:] = [d for d in dirs if d != "__MACOSX"]
                for file in files:
                    if file == ".DS_Store":
                        continue
                    full_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full_path, src_dir)
                    zf.write(full_path, arcname)
        os.replace(tmp_zip, output_path)
