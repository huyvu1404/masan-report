"""
Convert CMS buzz records (from fetch_data) to the DataFrame format
expected by process_data.

Column mapping:
    CMS field       → DataFrame column
    id              → Id
    (topic_name)    → Topic
    title           → Title
    content         → Content
    description     → Description
    url             → UrlComment, UrlTopic
    published_at    → PublishedDate
    sentiment       → Sentiment   (e.g. "Positive", "Negative", "Neutral", "None")
    site_name       → SiteName
    site_id         → SiteId
    channel         → Channel     (normalised to CHANNEL_ORDER values)
    author_name     → Author
    author_id       → AuthorId
    parent_id       → ParentId
    type            → Type        (e.g. "TIKTOK_COMMENT" — contains "topic"/"comment")
    labels[0..3]    → Labels1..4
"""

from __future__ import annotations

import pandas as pd

# Normalise CMS channel names → values used in CHANNEL_ORDER / process_data
_CHANNEL_NORM: dict[str, str] = {
    "Facebook Page":  "Facebook",
    "Facebook Group": "Facebook",
    "Facebook User":  "Facebook",
    "News":           "News",
    "Forum":          "Forum",
    "TikTok":         "Tiktok",
    "LinkedIn":       "Linkedin",
    "SNS":            "Social",
    "Threads":        "Threads",
    "YouTube":        "Youtube",
    "E-commerce":     "E-commerce",
    "Review":         "Review",
}

_EMPTY_COLS = [
    "Id", "Topic", "Title", "Content", "Description",
    "UrlComment", "UrlTopic", "PublishedDate", "Sentiment",
    "SiteName", "SiteId", "Channel", "Author", "AuthorId",
    "ParentId", "Type", "Labels1", "Labels2", "Labels3", "Labels4",
    "Level", "Tags",
]


def buzzes_to_dataframe(records: list[dict]) -> pd.DataFrame:
    """
    Convert CMS buzz records (từ fetch_data) sang DataFrame cho process_data.
    Mỗi record đã có field `topic` được gắn sẵn từ fetch_data.

    Returns:
        DataFrame ready for process_data(df, ...).
    """
    rows: list[dict] = []

    for rec in records:
        labels: list[str] = rec.get("labels") or []
        raw_channel: str = rec.get("channel") or ""
        rows.append({
            "Id":          rec.get("id"),
            "Topic":       rec.get("topic") or "",
            "Title":       rec.get("title") or "",
            "Content":     rec.get("content") or "",
            "Description": rec.get("description") or "",
            "UrlComment":  rec.get("url") or "",
            "UrlTopic":    rec.get("url") or "",
            "PublishedDate": rec.get("published_at"),
            "Sentiment":   (rec.get("sentiment") or "None").capitalize(),
            "SiteName":    rec.get("site_name") or "",
            "SiteId":      rec.get("site_id") or "",
            "Channel":     _CHANNEL_NORM.get(raw_channel, raw_channel),
            "Author":      rec.get("author_name") or "",
            "AuthorId":    rec.get("author_id") or "",
            "ParentId":    rec.get("parent_id") or rec.get("id"),
            "Type":        rec.get("type") or "",
            "Labels1":     labels[0] if len(labels) > 0 else None,
            "Labels2":     labels[1] if len(labels) > 1 else None,
            "Labels3":     labels[2] if len(labels) > 2 else None,
            "Labels4":     labels[3] if len(labels) > 3 else None,
            "Level":       None,
            "Tags":        None,
            })

    if not rows:
        return pd.DataFrame(columns=_EMPTY_COLS)

    return pd.DataFrame(rows)
