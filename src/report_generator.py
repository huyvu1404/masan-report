
from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Any, Optional

from src.chart_editor import PptxChartEditor
from src.text_editor import PptxTextEditor
from src.llm_client import generate_slide_texts

# Slide numbers in Template_full_fixed.pptx
# Slides 1,2,4,7 are fixed (cover + section dividers) — data slides only:
_S1 = 3  # Overview (Tổng quan)
_S2 = 5  # Competitive landscape
_S3 = 6  # Competitor health & channels
_S4 = 8  # Brand analysis
_S5 = 9  # Top posts table

# Direct mapping: process_data chart key → chart filename in template
CHART_FILE_MAP: dict[str, str] = {
    "weekly_bar":              "chart1.xml",
    "channel_doughnut":        "chart2.xml",
    "comp_discussion_bar":     "chart3.xml",
    "brand_sentiment_bar":     "chart4.xml",
    "main_sentiment_bar":      "chart5.xml",
    "sov_doughnut":            "chart6.xml",
    "daily_line":              "chart7.xml",
    "brand_discussion_bar":    "chart8.xml",
    # chart9-12: per-competitor topic health (dynamic, see below)
    "comp_sentiment_bar":      "chart13.xml",
    # chart14-17: per-competitor sources (dynamic, see below)
    "comp_channel_bar":        "chart18.xml",
    # "verbatim_bar":            "chart19.xml",
    "topic_health":            "chart20.xml",   # main brand topic health (slide 4)
    "main_sentiment_doughnut": "chart21.xml",
    "main_sources_bar":        "chart22.xml",
    "main_daily_line":         "chart23.xml",
}

# Per-competitor chart files ordered by competitor index (0-based)
COMP_TOPIC_HEALTH_FILES = ["chart9.xml", "chart10.xml", "chart11.xml", "chart12.xml"]
COMP_SOURCE_FILES       = ["chart14.xml", "chart15.xml", "chart16.xml", "chart17.xml"]

# Charts that must always be pre-cleared so template data never bleeds through
# when there is no new data for that chart.
ALWAYS_CLEAR_CHARTS = {"channel_doughnut", "topic_health", "main_sources_bar", "main_daily_line"}



# ── Format conversion ─────────────────────────────────────────────────────────

def _to_editor_series(series_list: list[dict]) -> list[dict]:
    """Convert process_data series format to PptxChartEditor.update_all_series format."""
    return [
        {
            "series_index": s["index"],
            "series_name":  s.get("title") or None,
            "categories":   s.get("cats") or None,
            "values":       s.get("vals") or None,
        }
        for s in series_list
    ]


# ── Verbatim truncation ───────────────────────────────────────────────────────

def _truncate_vb(brand: str, posts: list, max_post: int = 3, max_words: int = 20) -> str:
    truncated = []
    for post in posts[:max_post]:
        parts: list[str] = []
        for col in ["Title", "Content", "Description"]:
            val = (post.get(col) or "").strip()
            if val and val not in parts:
                parts.append(val)
        text = " ".join(" - ".join(parts).split()[:max_words])
        truncated.append(text)
    return "\n".join(f"{content} - {brand}" for content in truncated)


# ── Chart population ──────────────────────────────────────────────────────────

def _write_per_comp_charts(
    editor: PptxChartEditor,
    charts: dict,
    comps: list[str],
    key_prefix: str,
    files: list[str],
) -> None:
    for idx, comp in enumerate(comps[:len(files)]):
        key      = f"{key_prefix}_{comp}"
        filename = files[idx]
        try:
            editor.clear_series_beyond(filename, 0)
        except Exception:
            pass
        if key not in charts or not charts[key]["series"]:
            continue
        try:
            editor.update_all_series(filename, _to_editor_series(charts[key]["series"]))
        except Exception as exc:
            print(f"⚠️  Chart {filename} ({key}): {exc}")
    for idx in range(len(comps), len(files)):
        try:
            editor.clear_series_beyond(files[idx], 0)
        except Exception:
            pass


def _write_charts(editor: PptxChartEditor, data: dict) -> None:
    charts     = data["charts"]
    comps      = data["competitors"]
    all_brands = data["all_brands"]

    # Static chart key → file mapping
    for chart_key, filename in CHART_FILE_MAP.items():
        if chart_key in ALWAYS_CLEAR_CHARTS:
            try:
                editor.clear_series_beyond(filename, 0)
            except Exception:
                pass
        if chart_key not in charts or not charts[chart_key]["series"]:
            continue
        try:
            editor.update_all_series(filename, _to_editor_series(charts[chart_key]["series"]))
        except Exception as exc:
            print(f"⚠️  Chart {filename} ({chart_key}): {exc}")

    # chart7 (daily_line) has one series per brand — clear any extra template series
    n_daily_series = len(charts.get("daily_line", {}).get("series", []))
    try:
        editor.clear_series_beyond("chart7.xml", n_daily_series)
    except Exception as exc:
        print(f"⚠️  clear_series_beyond chart7.xml: {exc}")

    # Per-competitor topic health (chart9–12) and sources (chart14–17)
    _write_per_comp_charts(editor, charts, comps, "comp_topic_health", COMP_TOPIC_HEALTH_FILES)
    _write_per_comp_charts(editor, charts, comps, "comp_sources",      COMP_SOURCE_FILES)


# ── Text & table population ───────────────────────────────────────────────────

def _write_texts(editor: PptxTextEditor, data: dict, llm_texts: dict) -> None:
    brand   = data["main_brand"]
    period  = data["period"]
    metrics = data["all_brand_metrics"][brand]
    comps   = data["competitors"]
    posts   = data["top_posts"]
    topics  = data["top_topics"]
    vb      = data["verbatims"]
    bm      = data["brand_metrics"]
    total   = data["total_current"]

    # ── Slide _S1: Overview ──
    editor.replace_by_shape(_S1, "TextBox 10",
                            f"{metrics['total']:,}\nTHẢO LUẬN")

    top_topic_str = topics[0] if topics else ""

    overview_body = (
        f"Từ {period['start']} đến {period['end']}, {brand} ghi nhận "
        f"{total:,} thảo luận"
        + (f" đa số tin đề cập về {top_topic_str}." if top_topic_str else ".")
        + f"\n\nThảo luận tích cực chiếm {round(metrics['pos_rate']*100,1)}%, "
          f"tiêu cực chiếm {round(metrics['neg_rate']*100,1)}%."
    )
    overview_body = re.sub(r'^Tổng\s+quan\s*[:\-–]?\s*', '', overview_body, flags=re.IGNORECASE).strip()
    overview = f"Tổng quan:\n{overview_body}"
    editor.replace_by_shape(_S1, "Rectangle 15", overview, bold_first_line=True)

    conclusion = llm_texts.get("slide1_conclusion") or (
        f"{brand} ghi nhận NSR {metrics['nsr']}% trong kỳ báo cáo. "
        "Cần theo dõi các chủ đề nổi bật và duy trì tương tác tích cực."
    )
    editor.replace_by_shape(_S1, "Rectangle 13", conclusion)

    channels = bm[brand]["channels"]
    if channels:
        top_ch  = max(channels.items(), key=lambda x: x[1])
        ch_text = (f"Nền tảng:\nChiếm {round(top_ch[1]*100)}% tổng thảo luận, "
                   f"{top_ch[0]} là kênh truyền thông thu hút nhiều thảo luận nhất")
    else:
        ch_text = "Nền tảng:\nChưa có dữ liệu kênh"
    editor.replace_by_shape(_S1, "Rectangle 18", ch_text)

    if posts:
        p0 = posts[0]
        post_text = next(
            (" ".join((p0.get(f) or "").split()[:22]) for f in ("Title", "Content", "Description") if (p0.get(f) or "").strip()),
            "(Không có bài đăng nổi bật được ghi nhận)"
        )
    else:
        post_text = "(Không có bài đăng nổi bật được ghi nhận)"

    editor.replace_by_shape(_S1, "Rectangle 20", f"Bài đăng nổi bật:\n{post_text}")

    comp_bullets = []
    for comp in comps[:4]:
        comp_posts = bm.get(comp, {}).get("top_posts", [])
        snippet = ""
        if comp_posts:
            p = comp_posts[0]
            for field in ("Title", "Content", "Description"):
                text = (p.get(field) or "").strip()
                if text:
                    snippet = " ".join(text.split()[:10])
                    break
        comp_bullets.append(f"{comp}: {snippet or '(Không có bài đăng nổi bật)'}")

    comp_body = "\n".join(comp_bullets) if comp_bullets else "(Không có bài đăng nổi bật được ghi nhận)"
    editor.replace_by_shape(_S1, "Rectangle 22", f"Bài đăng nổi bật:\n{comp_body}", bold_first_line=True)

    # ── Slide _S2: Competitive ──
    editor.replace_by_shape(_S2, "TextBox 5",
                            f"{data['total_all']:,}\nTHẢO LUẬN")
    # object 3: "TỔNG QUAN THẢO LUẬN\nMasan ... 46,634 thảo luận."
    # Para 1, Run 1 chứa số tổng thảo luận — update đúng run, giữ nguyên font
    editor.replace_run_in_shape(_S2, "object 3", 1, 1, f"{data['total_all']:,}")
    editor.replace_by_shape(_S2, "Rounded Rectangle 2",
                            f"{period['start_label']} – {period['end_label']}")

    competitive_text = llm_texts.get("slide2_competitive")
    if not competitive_text:
        bullets = []
        for comp in comps[:5]:
            cm    = bm.get(comp, {})
            total = cm.get("current", {}).get("total", 0)
            srcs  = cm.get("sources", [])
            src   = srcs[0]["source"] if srcs else ""
            bullets.append(f"{comp}: {total:,} thảo luận. Nguồn chính: {src}.")
        competitive_text = "\n".join(bullets)
    editor.replace_by_shape(_S2, "TextBox 16", competitive_text)

    # ── Slide _S3: Competitor health ──
    comp_shapes = [
        "Rectangle 37", "Rectangle 40", "Rectangle 43", "Rectangle 46",
        "Rectangle 51", "Rectangle 57", "Rectangle 60", "Rectangle 63",
    ]

    blanks = len(comp_shapes) / 2 - len(comps)
    for i, comp in enumerate(comps[:4]):
        editor.replace_by_shape(_S3, comp_shapes[i], comp.upper())
        editor.replace_by_shape(_S3, comp_shapes[i+4], comp.upper())
        if blanks > 0:
            editor.replace_by_shape(_S3, comp_shapes[i + int(blanks)], "")
            editor.replace_by_shape(_S3, comp_shapes[i + int(blanks) + 4], "")

    # ── Slide _S4: Brand analysis ──
    editor.replace_by_shape(_S4, "TextBox 10", f"NSR: {metrics['nsr']}%")

    analysis = llm_texts.get("slide4_analysis") or (
        f"Trong kỳ {period['start_label']} – {period['end_label']}, "
        f"{brand} ghi nhận {metrics['total']:,} thảo luận với NSR {metrics['nsr']}%. "
        f"Tích cực chiếm {round(metrics['pos_rate']*100,1)}%, "
        f"tiêu cực {round(metrics['neg_rate']*100,1)}%."
    )
    editor.replace_by_shape(_S4, "TextBox 4", analysis)
    pos_text = _truncate_vb(brand, vb["pos"])
    neg_text = _truncate_vb(brand, vb["neg"])
    editor.replace_by_shape(_S4, "Rectangle 5",
                            pos_text or "(Chưa có dữ liệu tích cực)")
    editor.replace_by_shape(_S4, "Rectangle 10",
                            neg_text or "(Chưa có dữ liệu tiêu cực)")

    # ── Slide _S5: Top posts table ──
    editor.replace_by_shape(_S5, "object 3", "CÁC BÀI ĐĂNG NỔI BẬT")
    editor.fill_table_on_slide(_S5, posts)


def _write_hyperlinks(editor: PptxTextEditor, data: dict, llm_texts: dict) -> None:
    """Attach hyperlinks to featured-post shapes and the slide-5 table after text is set."""
    posts     = data["top_posts"]
    bm        = data["brand_metrics"]
    comps     = data["competitors"]
    all_brands = data["all_brands"]

    def _post_url(brand: str) -> str:
        ps = bm.get(brand, {}).get("top_posts", [])
        return ps[0].get("UrlTopic", "") if ps else ""

    # ── Slide _S1: Rectangle 20 — Masan Group featured post (full-content hyperlink) ──
    if posts:
        url = posts[0].get("UrlTopic", "")
        if url and url not in ("None", "nan"):
            editor.add_shape_para_hyperlinks(_S1, "Rectangle 20", {1: url})

    # ── Slide _S1: Rectangle 22 — competitors: hyperlink snippet, not brand name ──
    comp_urls = {}
    for i, comp in enumerate(comps[:4]):
        url = _post_url(comp)
        if url and url not in ("None", "nan"):
            comp_urls[i + 1] = url  # runnable para 0 = header, 1+ = competitor lines
    if comp_urls:
        editor.hyperlink_para_content(_S1, "Rectangle 22", comp_urls)

    # ── Slide _S2: TextBox 16 — append " (URL)" at end of each brand paragraph ──
    competitive_text = llm_texts.get("slide2_competitive")
    brand_urls: dict[int, str] = {}
    if competitive_text:
        # LLM reorders brands (puts top brand first) — detect brand by name in each line
        all_brand_urls = {b: _post_url(b) for b in all_brands}
        all_brand_urls = {b: u for b, u in all_brand_urls.items() if u and u not in ("None", "nan")}
        for i, line in enumerate(competitive_text.split("\n")):
            line_lower = line.lower()
            for brand, url in all_brand_urls.items():
                if brand.lower() in line_lower:
                    brand_urls[i] = url
                    break
    else:
        ### 
        for i, comp in enumerate(comps[:5]):
            u = _post_url(comp)
            if u and u not in ("None", "nan"):
                brand_urls[i] = u
    print(f"🔗 TextBox 16 brand_urls: {len(brand_urls)}/{len(all_brands)} brands have URL")
    for k, v in brand_urls.items():
        print(f"   [{k}]: {v[:60]}")
    if brand_urls:
        editor.append_hyperlink_run_to_shape(_S2, "TextBox 16", brand_urls)

    # ── Slide _S5: table column 1 (CHỦ ĐỀ) ──
    for i, post in enumerate(posts):
        url = post.get("UrlTopic", "")
        if url and url not in ("None", "nan"):
            editor.add_table_cell_hyperlink(_S5, i, 1, url)


# ── Main entry point ──────────────────────────────────────────────────────────

def generate_report(
    data: dict[str, Any],
    template_path: Path,
    api_key: Optional[str] = None,
) -> bytes:
    """
    Populate PPTX template from process_data() output.

    Three-pass approach:
      Pass 0 — LLM text generation (skipped if api_key is None)
      Pass 1 — chart data via PptxChartEditor (lxml, correct namespace handling)
      Pass 2 — text/table content via PptxTextEditor (lxml)

    Returns raw bytes of the generated .pptx file.
    """
    # Pass 0: generate slide texts with LLM (optional)
    llm_texts: dict = {}
    if api_key:
        print("🤖 Generating slide texts with LLM...")
        try:
            llm_texts = generate_slide_texts(data, api_key)
        except Exception as exc:
            print(f"⚠️  LLM generation failed, falling back to hardcoded text: {exc}")

    with tempfile.TemporaryDirectory(prefix="pptx_gen_") as tmp_dir:
        pass1_path = os.path.join(tmp_dir, "pass1.pptx")
        pass2_path = os.path.join(tmp_dir, "pass2.pptx")

        # Pass 1: edit charts
        chart_editor = PptxChartEditor(str(template_path))
        _write_charts(chart_editor, data)
        chart_editor.save(pass1_path)

        # Pass 2: edit text & tables
        text_editor = PptxTextEditor(pass1_path)
        _write_texts(text_editor, data, llm_texts)
        _write_hyperlinks(text_editor, data, llm_texts)
        text_editor.save(pass2_path)

        return Path(pass2_path).read_bytes()
