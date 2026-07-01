"""
process_data.py
===============
Process raw social listening DataFrame into structured data for charts and LLM text.

Usage:
    from src.process_data import process_data

    result = process_data(
        df=df,
        main_brand="Masan Consumer",
        competitors=["Brand A", "Brand B"],
        start_date=datetime(2025, 6, 1),
        end_date=datetime(2025, 6, 7),
    )

    # chart series ready for _set_chart_series_data / PptxChartEditor
    result["charts"]["topic_health"]["series"]

    # rich dict for building LLM prompts
    result["llm_context"]
"""

from __future__ import annotations

import pandas as pd
from datetime import date, datetime, timedelta
from typing import Any


# ──────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ──────────────────────────────────────────────────────────────────────────────

def _sentiment_metrics(frame: pd.DataFrame) -> dict:
    """Return pos/neg/neu/total/nsr/rates dict for a frame."""
    if frame.empty:
        return {"pos": 0, "neg": 0, "neu": 0, "total": 0, "nsr": 0,
                "pos_rate": 0.0, "neg_rate": 0.0, "neu_rate": 0.0}
    s = frame["Sentiment"].str.lower().value_counts()
    pos = int(s.get("positive", 0))
    neg = int(s.get("negative", 0))
    neu = int(s.get("neutral", 0))
    total = pos + neg + neu
    nsr = round((pos - neg) / (pos + neg) * 100) if (pos + neg) > 0 else 0
    denom = total or 1
    return {
        "pos": pos, "neg": neg, "neu": neu, "total": total, "nsr": nsr,
        "pos_rate": round(pos / denom, 6),
        "neg_rate": round(neg / denom, 6),
        "neu_rate": round(neu / denom, 6),
    }


def _channel_dist(frame: pd.DataFrame) -> dict[str, float]:
    if frame.empty or "Channel" not in frame.columns:
        return {}
    counts = frame["Channel"].dropna().value_counts()
    total = counts.sum() or 1
    return {ch: round(int(cnt) / total, 6) for ch, cnt in counts.items()}


def _top_sources(frame: pd.DataFrame, n: int = 5) -> list[dict]:
    if frame.empty or "SiteName" not in frame.columns:
        return []
    counts = frame["SiteName"].fillna("Unknown").value_counts().head(n)
    return [{"source": str(s), "count": int(c)} for s, c in counts.items()]


def _top_sources_by_engagement(frame: pd.DataFrame, n: int = 6) -> list[dict]:
    """Top N sources sorted by sum of Interaction column (falls back to mention count)."""
    if frame.empty or "SiteName" not in frame.columns:
        return []
    if "Interaction" in frame.columns:
        tmp = frame.copy()
        tmp["Interaction"] = pd.to_numeric(tmp["Interaction"], errors="coerce").fillna(0)
        agg = tmp.groupby("SiteName")["Interaction"].sum().nlargest(n)
    else:
        agg = frame["SiteName"].fillna("Unknown").value_counts().head(n)
    return [{"source": str(s), "count": int(c)} for s, c in agg.items()]


def _daily_counts(frame: pd.DataFrame) -> dict[str, int]:
    if frame.empty:
        return {}
    g = frame.groupby(frame["PublishedDate"].dt.date).size()
    return {str(d): int(c) for d, c in g.items()}

def _topic_health(frame: pd.DataFrame, label_cols: list[str], top_n: int = 10) -> dict[str, dict]:
    """
    Aggregate sentiment across all Labels1–4 columns.

    Each row can contribute to multiple topics (one per non-empty label column).
    Returns {topic: {pos, neg, neu, total, nsr, pos_rate, neg_rate, neu_rate}}.
    """
    available = [c for c in label_cols if c in frame.columns]
    if not available:
        return {}

    parts = []
    for col in available:
        sub = frame[[col, "Sentiment"]].copy()
        sub = sub.rename(columns={col: "_label"})
        sub = sub.dropna(subset=["_label"])
        sub = sub[sub["_label"].astype(str).str.strip() != ""]
        sub["_label"] = sub["_label"].astype(str).str.strip()
        parts.append(sub)

    if not parts:
        return {}

    combined = pd.concat(parts, ignore_index=True)
    combined["_sent"] = combined["Sentiment"].str.lower()

    grp = combined.groupby(["_label", "_sent"]).size().unstack(fill_value=0)
    grp = grp.reindex(columns=["positive", "negative", "neutral"], fill_value=0)
    grp["total"] = grp.sum(axis=1)
    grp = grp.sort_values("total", ascending=False).head(top_n)

    result: dict[str, dict] = {}
    for label, row in grp.iterrows():
        pos, neg, neu, total = int(row["positive"]), int(row["negative"]), int(row["neutral"]), int(row["total"])
        nsr = round((pos - neg) / (pos + neg) * 100) if (pos + neg) > 0 else 0
        denom = total or 1
        result[str(label)] = {
            "pos": pos, "neg": neg, "neu": neu, "total": total, "nsr": nsr,
            "pos_rate": round(pos / denom, 6),
            "neg_rate": round(neg / denom, 6),
            "neu_rate": round(neu / denom, 6),
        }
    return result


def _buzz_types(frame: pd.DataFrame) -> dict[str, dict]:
    """Aggregate sentiment breakdown by Type column."""
    if frame.empty or "Type" not in frame.columns:
        return {}
    grp = (frame.groupby(["Type", frame["Sentiment"].str.lower()])
           .size().unstack(fill_value=0))
    grp = grp.reindex(columns=["positive", "negative", "neutral"], fill_value=0)
    grp["total"] = grp.sum(axis=1)
    grp = grp.sort_values("total", ascending=False)
    return {
        str(t): {
            "pos": int(r["positive"]), "neg": int(r["negative"]),
            "neu": int(r["neutral"]), "total": int(r["total"]),
        }
        for t, r in grp.iterrows()
    }


def _level_dist(frame: pd.DataFrame) -> dict[str, int]:
    if frame.empty or "Level" not in frame.columns:
        return {}
    counts = frame["Level"].dropna().value_counts()
    return {str(lv): int(cnt) for lv, cnt in counts.items()}


def _tags_dist(frame: pd.DataFrame) -> dict[str, int]:
    if frame.empty or "Tags" not in frame.columns:
        return {}
    counts = frame["Tags"].dropna().value_counts()
    return {str(tag): int(cnt) for tag, cnt in counts.items()}


def _top_posts(
    frame: pd.DataFrame,
    n: int = 10,
    sentiment: str | None = None,
) -> list[dict]:
    if frame.empty:
        return []

    df = frame.copy()

    if sentiment is not None:
        df = df[df["Sentiment"].str.lower() == sentiment.lower()]

    if df.empty:
        return []

    posts = []

    for _, group in df.groupby("ParentId"):
        topic_mask = (
            group["Type"]
            .fillna("")
            .str.lower()
            .str.contains("topic")
        )

        if topic_mask.any():
            topic_row = group[topic_mask].iloc[0]
            comments  = group.drop(index=topic_row.name)
        else:
            # Comments-only group: Title on every row = parent post title
            topic_row = group.iloc[0]
            comments  = group
        comment_contents = [
            f"[{row.get('Sentiment', 'Neutral')}] - {str(row.get('Content', '') or '')}"
            for _, row in comments.iterrows()
            if str(row.get("Content", "")).strip()
        ]
        posts.append({
            "Rank": 0,
            "Id": str(topic_row.get("Id", "")),
            "Title": str(topic_row.get("Title", "") or ""),
            "Content": str(topic_row.get("Content", "") or ""),
            "Description": str(topic_row.get("Description", "") or ""),
            "SiteName": str(topic_row.get("SiteName", "")),
            "UrlTopic": str(topic_row.get("UrlTopic", "")),
            "Channel": str(topic_row.get("Channel", "")),
            "Author": str(topic_row.get("Author", "")),
            "Sentiment": str(topic_row.get("Sentiment", "Neutral")),
            "Interaction": len(comments),
            "Comment": comment_contents,
            "PublishedDate": topic_row.get("PublishedDate"),
        })

    # Sort theo số lượng comment giảm dần, rồi theo ngày
    posts = sorted(
        posts,
        key=lambda x: (
            x["Interaction"]
        ),
        reverse=True,
    )[:n]

    for i, post in enumerate(posts, start=1):
        post["Rank"] = i

    return posts


# ──────────────────────────────────────────────────────────────────────────────
# Chart series builders
# Each returns a list of series dicts compatible with _set_chart_series_data
# ──────────────────────────────────────────────────────────────────────────────

def _series(title: str, cats: list, vals: list, idx: int = 0,
            show_val: bool = True, show_pct: bool = False,
            fill_color: str = "") -> dict:
    return {
        "index": idx, "order": idx, "title": title,
        "cats": [str(c) for c in cats],
        "vals": [str(v) for v in vals],
        "data_labels": {
            "show_val": show_val,
            "show_percent": show_pct,
            "show_cat_name": False,
            "show_ser_name": False,
        },
        "style": {"fill_color": fill_color} if fill_color else {},
    }


def _build_weekly_bar(main_brand: str, curr_total: int, prev_total: int,
                      curr_label: str, prev_label: str) -> list[dict]:
    return [_series(main_brand, [prev_label, curr_label],
                    [prev_total, curr_total], show_val=True)]


def _build_channel_doughnut(channels: dict[str, float]) -> list[dict]:
    sorted_ch = sorted(channels.items(), key=lambda x: x[1], reverse=True)
    return [_series("Channel", [c[0] for c in sorted_ch],
                    [round(c[1], 6) for c in sorted_ch],
                    show_val=False, show_pct=True)]


def _build_comp_discussion_bar(competitors: list[str], comp_weekly: dict,
                                prev_label: str, curr_label: str) -> list[dict]:
    prev_vals = [comp_weekly.get(c, {}).get("prev", 0) for c in competitors]
    curr_vals = [comp_weekly.get(c, {}).get("curr", 0) for c in competitors]
    return [
        _series(prev_label, competitors, prev_vals, idx=0),
        _series(curr_label, competitors, curr_vals, idx=1, fill_color="accent2"),
    ]


def _build_brand_sentiment_bar(all_brands: list[str],
                                all_brand_metrics: dict[str, dict]) -> list[dict]:
    pos = [all_brand_metrics.get(b, {}).get("pos_rate", 0) for b in all_brands]
    neu = [all_brand_metrics.get(b, {}).get("neu_rate", 0) for b in all_brands]
    neg = [all_brand_metrics.get(b, {}).get("neg_rate", 0) for b in all_brands]
    nsr = [round(all_brand_metrics.get(b, {}).get("nsr", 0) / 100, 6) for b in all_brands]
    return [
        _series("Positive",  all_brands, pos, idx=0, fill_color="00B050"),
        _series("Neutral",   all_brands, neu, idx=1, fill_color="bg1"),
        _series("Negative",  all_brands, neg, idx=2, fill_color="FF0000"),
        _series("NSR",       all_brands, nsr, idx=3, fill_color="0070C0"),
    ]

# Fixed order matching chart18 template series indices (0-8)
CHANNEL_ORDER = ["Facebook", "News", "Forum", "Tiktok", "Linkedin",
                 "Social", "Threads", "Youtube", "E-commerce"]

def _build_compt_channel_bar(brands: list[str], comp_channels: dict[str, dict[str, float]]) -> list[dict]:
    """Build series for chart18: one series per channel in template order, categories = brands."""
    series_list = []
    for idx, ch in enumerate(CHANNEL_ORDER):
        vals = [round(comp_channels.get(brand, {}).get(ch, 0), 6) for brand in brands]
        series_list.append(_series(ch, brands, vals, idx=idx))
    return series_list

def _build_sov_doughnut(all_brands: list[str],
                         all_brand_metrics: dict[str, dict]) -> list[dict]:
    totals = [all_brand_metrics.get(b, {}).get("total", 0) for b in all_brands]
    return [_series("Share of Voice", all_brands, totals,
                    show_val=False, show_pct=True)]


def _date_to_excel_serial(date_str: str) -> int:
    """Convert YYYY-MM-DD string to Excel serial date number (for dateAx charts)."""
    d = date.fromisoformat(date_str)
    return (d - date(1899, 12, 30)).days


def _build_daily_line(brand_daily_map: dict[str, dict[str, int]]) -> list[dict]:
    """Multi-brand line chart using DD/MM string categories (catAx, e.g. chart7)."""
    colors = ["accent1", "accent4", "FF237B", "AED8AF", "accent5"]
    series = []
    for idx, (brand, daily) in enumerate(brand_daily_map.items()):
        if not daily:
            continue
        sorted_days = sorted(daily.items())
        cats = [date.fromisoformat(d).strftime("%m/%d/%Y") for d, _ in sorted_days]
        vals = [c for _, c in sorted_days]
        series.append(_series(brand, cats, vals,
                               idx=idx, fill_color=colors[idx % len(colors)]))
    return series


def _build_main_daily_line(brand: str, daily: dict[str, int]) -> list[dict]:
    """Single-brand line chart using Excel serial date categories (dateAx, e.g. chart23)."""
    if not daily:
        return []
    sorted_days = sorted(daily.items())
    cats = [_date_to_excel_serial(d) for d, _ in sorted_days]
    vals = [c for _, c in sorted_days]
    return [_series(brand, cats, vals)]


def _build_topic_health_bar(topic_health: dict[str, dict], top_n: int = 6, brand_total: int = 0) -> list[dict]:
    """Series for CHỈ SỐ SỨC KHOẺ THEO CHỦ ĐỀ chart.
    Values are normalized by brand_total so they represent share of total brand mentions.
    4th series is Grand Total (pos+neu+neg for that topic / brand_total).
    """
    topics = list(topic_health.keys())[:top_n]
    denom = brand_total or 1
    pos   = [round(topic_health[t]["pos"] / denom, 6) for t in topics]
    neu   = [round(topic_health[t]["neu"] / denom, 6) for t in topics]
    neg   = [round(topic_health[t]["neg"] / denom, 6) for t in topics]
    grand = [round(topic_health[t]["total"] / denom, 6) for t in topics]
    return [
        _series("Positive",    topics, pos,   idx=0, fill_color="00B050"),
        _series("Neutral",     topics, neu,   idx=1, fill_color="tx1"),
        _series("Negative",    topics, neg,   idx=2, fill_color="C00000"),
        _series("Grand Total", topics, grand, idx=3, show_val=True),
    ]


def _build_sentiment_doughnut(metrics: dict) -> list[dict]:
    cats = ["Positive", "Neutral", "Negative"]
    vals = [metrics["pos_rate"], metrics["neu_rate"], metrics["neg_rate"]]
    return [_series("Sentiment", cats, vals, show_val=True, show_pct=False)]


def _build_sources_bar(sources: list[dict]) -> list[dict]:
    if not sources:
        return []
    return [_series("Total",
                    [s["source"] for s in sources],
                    [s["count"] for s in sources])]


def _build_verbatim_bar(metrics: dict) -> list[dict]:
    cats = ["Negative post", "Positive post"]
    return [
        _series("Share", cats, [metrics["neg"], metrics["pos"]], idx=0),
        _series("Gốc",   cats, [metrics["neg"], metrics["pos"]], idx=1),
    ]


# ──────────────────────────────────────────────────────────────────────────────
# LLM context builder
# ──────────────────────────────────────────────────────────────────────────────

def _build_trend_line(daily: dict[str, int]) -> str:
    """One-sentence description of the daily discussion trend."""
    if not daily:
        return "Không có dữ liệu xu hướng."
    sorted_days = sorted(daily.items())
    total = sum(v for _, v in sorted_days)
    avg   = round(total / len(sorted_days), 1)
    peak_day, peak_val = max(sorted_days, key=lambda x: x[1])
    # Simple trend: compare last 3 days to first 3 days
    first3 = [v for _, v in sorted_days[:3]]
    last3  = [v for _, v in sorted_days[-3:]]
    avg_first = sum(first3) / len(first3)
    avg_last  = sum(last3)  / len(last3)
    if avg_last > avg_first * 1.1:
        trend = "tăng"
    elif avg_last < avg_first * 0.9:
        trend = "giảm"
    else:
        trend = "ổn định"
    return (
        f"Trung bình {avg:,} thảo luận/ngày, xu hướng {trend}. "
        f"Đỉnh cao nhất vào {peak_day} với {peak_val:,} thảo luận."
    )


def _build_llm_context(main_brand: str, period: dict,
                        main_metrics: dict, prev_metrics: dict,
                        channels: dict, sources: list[dict],
                        topic_health: dict, buzz_types: dict,
                        all_brand_metrics: dict[str, dict],
                        competitors: list[str],
                        top_posts: list[dict],
                        total_all: int,
                        pos_verbatims: list[str] | None = None,
                        neg_verbatims: list[str] | None = None,
                        daily: dict[str, int] | None = None) -> dict:
    top_channel = max(channels.items(), key=lambda x: x[1]) if channels else ("N/A", 0)
    top_topics = list(topic_health.keys())[:5]
    neg_topics = sorted(topic_health.items(),
                        key=lambda x: x[1]["neg_rate"], reverse=True)[:3]
    change_pct = (
        round((main_metrics["total"] - prev_metrics["total"]) / prev_metrics["total"] * 100, 1)
        if prev_metrics["total"] > 0 else 0
    )
    comp_summary = []
    for comp in competitors:
        m = all_brand_metrics.get(comp, {})
        comp_summary.append({
            "brand": comp,
            "total": m.get("total", 0),
            "nsr": m.get("nsr", 0),
            "pos_pct": round(m.get("pos_rate", 0) * 100, 1),
        })

    # Pre-built narrative snippets for LLM prompts
    sentiment_line = (
        f"Tích cực: {round(main_metrics['pos_rate']*100,1)}% | "
        f"Trung lập: {round(main_metrics['neu_rate']*100,1)}% | "
        f"Tiêu cực: {round(main_metrics['neg_rate']*100,1)}% | "
        f"NSR: {main_metrics['nsr']}%"
    )
    comp_line = "; ".join(
        f"{c['brand']} ({c['total']:,} thảo luận, NSR {c['nsr']}%)"
        for c in comp_summary
    )
    topic_line = ", ".join(
        f"{t} (NSR {v['nsr']}%)" for t, v in list(topic_health.items())[:5]
    )
    buzz_line = ", ".join(
        f"{t}: {v['total']}" for t, v in list(buzz_types.items())[:5]
    ) if buzz_types else "N/A"

    data_trend_line = _build_trend_line(daily or {})


    return {
        # Key numbers
        "brand": main_brand,
        "period": f"{period['start_label']} – {period['end_label']}",
        "period_full": f"{period['start']} – {period['end']}",
        "total_current": main_metrics["total"],
        "total_previous": prev_metrics["total"],
        "change_pct": change_pct,
        "total_all_brands": total_all,

        # Sentiment
        "nsr": main_metrics["nsr"],
        "pos_pct": round(main_metrics["pos_rate"] * 100, 1),
        "neg_pct": round(main_metrics["neg_rate"] * 100, 1),
        "neu_pct": round(main_metrics["neu_rate"] * 100, 1),
        "sentiment_line": sentiment_line,

        # Previous period
        "prev_nsr": prev_metrics["nsr"],
        "prev_total": prev_metrics["total"],

        # Channels & sources
        "top_channel": top_channel[0],
        "top_channel_pct": round(top_channel[1] * 100, 1),
        "top_sources": [s["source"] for s in sources[:3]],

        # Topics
        "top_topics": top_topics,
        "negative_topics": [t for t, _ in neg_topics],
        "topic_line": topic_line,

        # Buzz types
        "buzz_types": buzz_types,
        "buzz_line": buzz_line,

        # Competitors
        "competitors_summary": comp_summary,
        "comp_line": comp_line,

        # Verbatims & trend (used in slide 4 LLM prompt)
        "pos_verbatims": pos_verbatims,
        "neg_verbatims": neg_verbatims,
        
        "data_trend_line": data_trend_line,
    }


# ──────────────────────────────────────────────────────────────────────────────
# Main entry point
# ──────────────────────────────────────────────────────────────────────────────

LABEL_COLS = ["Labels1", "Labels2", "Labels3", "Labels4"]


def process_data(
    df: pd.DataFrame,
    main_brand: str,
    competitors: list[str],
    start_date: datetime,
    end_date: datetime,
) -> dict[str, Any]:
    """
    Process raw social listening data into chart-ready and LLM-ready structures.

    Args:
        df:          Raw DataFrame with columns: Id, TopicId, Topic, Title,
                     Content, Description, UrlComment, UrlTopic, PublishedDate,
                     Sentiment, SiteName, SiteId, Channel, Author, AuthorId,
                     ParentId, Labels1–4, Type, Level, Tags.
        main_brand:  Primary brand to analyse.
        competitors: List of competitor brand names.
        start_date:  Start of current reporting period (inclusive).
        end_date:    End of current reporting period (inclusive).

    Returns:
        Dict with keys: period, main_brand, competitors, brand_metrics,
        all_brand_metrics, charts, llm_context, top_posts, verbatims, total_all.
    """
    df = df.copy()
    for col in ("Title", "Content", "Description"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    df["PublishedDate"] = pd.to_datetime(df["PublishedDate"], format="%Y-%m-%d %H:%M:%S", errors="coerce")

    # Normalize MSN tag → canonical brand names (always applied)
    if main_brand == "MSN":
        main_brand = "Masan Group"

    def _norm_topic(t: str) -> str:
        t = str(t).strip()
        if t == "MSN":
            return "Masan Group"
        if t.endswith(" - MSN"):
            return t[:-6].strip()
        return t

    df["Topic"] = df["Topic"].apply(_norm_topic)
    competitors = [_norm_topic(c) for c in competitors]
    all_brands = [main_brand] + competitors
    # ── Period slices ──
    mask_curr = (df["PublishedDate"] >= start_date) & (df["PublishedDate"] <= end_date)
    df_curr = df[mask_curr]
    df_main = df_curr[df_curr["Topic"] == main_brand]
    df_all_curr = df_curr[df_curr["Topic"].isin(all_brands)]

    days = (end_date.date() - start_date.date()).days + 1

    prev_end = start_date - timedelta(seconds=1)
    prev_start = start_date - timedelta(days=days)
    mask_prev  = (df["PublishedDate"] >= prev_start) & (df["PublishedDate"] <= prev_end)
    df_prev_all = df[mask_prev & df["Topic"].isin(all_brands)]
    df_main_prev = df[mask_prev & (df["Topic"] == main_brand)]

    prev_label = f"{prev_start.strftime('%d/%m')} – {prev_end.strftime('%d/%m')}"
    curr_label = f"{start_date.strftime('%d/%m')} – {end_date.strftime('%d/%m')}"

    period = {
        "start":       start_date.strftime("%d/%m/%Y"),
        "end":         end_date.strftime("%d/%m/%Y"),
        "start_label": start_date.strftime("%d/%m"),
        "end_label":   end_date.strftime("%d/%m"),
        "prev_label":  prev_label,
        "curr_label":  curr_label,
        "prev_start":  prev_start.strftime("%d/%m/%Y"),
        "prev_end":    prev_end.strftime("%d/%m/%Y"),
    }

    # ── Per-brand metrics (current period) ──
    all_brand_metrics: dict[str, dict] = {}
    for brand in all_brands:
        bdf = df_all_curr[df_all_curr["Topic"] == brand]
        all_brand_metrics[brand] = _sentiment_metrics(bdf)

    # ── Main brand full metrics ──
    main_metrics  = all_brand_metrics[main_brand]
    prev_metrics  = _sentiment_metrics(df_main_prev)
    main_channels = _channel_dist(df_main)
    main_sources  = _top_sources(df_main)
    main_daily    = _daily_counts(df_main)
    main_topic_h  = _topic_health(df_main, LABEL_COLS)
    main_buzz     = _buzz_types(df_main)
    main_levels   = _level_dist(df_main)
    main_tags     = _tags_dist(df_main)

    # ── Competitor full metrics ──
    brand_metrics: dict[str, dict] = {}
    for brand in all_brands:
        bdf      = df_all_curr[df_all_curr["Topic"] == brand]
        bdf_prev = df_prev_all[df_prev_all["Topic"] == brand]
        brand_metrics[brand] = {
            "current":     all_brand_metrics[brand],
            "previous":    _sentiment_metrics(bdf_prev),
            "channels":    _channel_dist(bdf),
            "sources":     _top_sources(bdf, 5),
            "top_posts":   _top_posts(bdf, 3),
            "daily":       _daily_counts(bdf),
            "topic_health": _topic_health(bdf, LABEL_COLS),
            "buzz_types":  _buzz_types(bdf),
            "levels":      _level_dist(bdf),
            "tags":        _tags_dist(bdf),
        }

    # ── Competitor weekly discussion comparison ──
    comp_weekly = {
        comp: {
            "prev": len(df_prev_all[df_prev_all["Topic"] == comp]),
            "curr": len(df_all_curr[df_all_curr["Topic"] == comp]),
        }
        for comp in competitors
    }

    posts  = _top_posts(df_main, n=5)                         # slide 5 table: top 5 by interaction
    pos_vb = _top_posts(df_main, n=3, sentiment="Positive")   # slide 4 verbatims + LLM slide 1 & 4
    neg_vb = _top_posts(df_main, n=3, sentiment="Negative")

    charts: dict[str, dict] = {
        # Slide 1
        "weekly_bar": {
            "series": _build_weekly_bar(
                main_brand, main_metrics["total"], prev_metrics["total"],
                curr_label, prev_label
            ),
        },
        "channel_doughnut": {
            "series": _build_channel_doughnut(main_channels),
        },
        "comp_discussion_bar": {
            "series": _build_comp_discussion_bar(
                competitors, comp_weekly, prev_label, curr_label
            ),
        },
        "brand_sentiment_bar": {
            "series": _build_brand_sentiment_bar(competitors, all_brand_metrics),
        },
        "main_sentiment_bar": {
            "series": [
                _series("Positive",  [main_brand], [main_metrics["pos_rate"]], idx=0, fill_color="00B050"),
                _series("Neutral",   [main_brand], [main_metrics["neu_rate"]], idx=1, fill_color="7F7F7F"),
                _series("Negative",  [main_brand], [main_metrics["neg_rate"]], idx=2, fill_color="FF0000"),
            ],
        },
        # Slide 2
        "sov_doughnut": {
            "series": _build_sov_doughnut(all_brands, all_brand_metrics),
        },
        "daily_line": {
            "series": _build_daily_line({b: brand_metrics[b]["daily"] for b in all_brands}),
        },
        "brand_discussion_bar": {
            "series": [
                _series(prev_label, all_brands,
                        [brand_metrics[b]["previous"]["total"] for b in all_brands], idx=0),
                _series(curr_label, all_brands,
                        [all_brand_metrics[b]["total"] for b in all_brands], idx=1,
                        fill_color="accent2"),
            ],
        },
        # Slide 3
        "comp_sentiment_bar": {
            "series": _build_brand_sentiment_bar(competitors, all_brand_metrics),
        },
        "comp_channel_bar": {
            "series": _build_compt_channel_bar(
                competitors, {brand: brand_metrics[brand]["channels"] for brand in competitors}
            ),
        },


        # Slide 4 – CHỈ SỐ SỨC KHOẺ THEO CHỦ ĐỀ
        "topic_health": {
            "series": _build_topic_health_bar(main_topic_h, brand_total=main_metrics["total"]),
            "raw": main_topic_h,
        },
        "main_sentiment_doughnut": {
            "series": _build_sentiment_doughnut(main_metrics),
        },
        "main_sources_bar": {
            "series": _build_sources_bar(_top_sources_by_engagement(df_main, 6)),
        },
        "main_daily_line": {
            "series": _build_main_daily_line(main_brand, main_daily),
        },
        "verbatim_bar": {
            "series": _build_verbatim_bar(main_metrics),
        },
    }
    for comp in competitors:
        charts[f"comp_sources_{comp}"] = {
            "series": _build_sources_bar(brand_metrics[comp]["sources"]),
        }
        charts[f"comp_topic_health_{comp}"] = {
            "series": _build_topic_health_bar(
                brand_metrics[comp]["topic_health"],
                brand_total=all_brand_metrics[comp]["total"],
            ),
            "raw": brand_metrics[comp]["topic_health"],
        }

    # Per-competitor individual charts (slide 3)

    # ── LLM context ──
    llm_context = _build_llm_context(
        main_brand, period, main_metrics, prev_metrics,
        main_channels, main_sources, main_topic_h, main_buzz,
        all_brand_metrics, competitors, posts, len(df_all_curr),
        pos_verbatims=pos_vb or "Không có bài đăng tích cực",
        neg_verbatims=neg_vb or "Không có bài đăng tiêu cực",
        daily=brand_metrics[main_brand]["daily"],
    )

    return {
        "period":            period,
        "main_brand":        main_brand,
        "competitors":       competitors,
        "all_brands":        all_brands,

        # Detailed per-brand data
        "brand_metrics":     brand_metrics,
        "all_brand_metrics": all_brand_metrics,      # current-period only, flat
        "comp_weekly":       comp_weekly,

        # Additional main-brand breakdowns
        "main_buzz_types":   main_buzz,
        "main_levels":       main_levels,
        "main_tags":         main_tags,

        # Pre-built chart series
        "charts":            charts,

        # LLM prompt data
        "llm_context":       llm_context,

        # Convenience shortcuts (mirrors keys inside llm_context)
        "top_topics":        llm_context["top_topics"],
        "total_current":     llm_context["total_current"],

        # Raw lists
        "top_posts":         posts,
        "verbatims":         {"pos": pos_vb, "neg": neg_vb},
        "total_all":         len(df_all_curr),
    }
