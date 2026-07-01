"""
Fetch and map buzz data from the CMS (GraphQL).

Required .env vars:
    CMS_URL   GraphQL endpoint
"""

from __future__ import annotations

import math
import os
from datetime import datetime, timedelta, timezone
from typing import Optional

import requests
from dotenv import load_dotenv

load_dotenv()

_CMS_URL = os.getenv("CMS_URL", "").rstrip("/")

_BASE_HEADERS = {
    "accept": "*/*",
    "accept-language": "en-US,en;q=0.9,vi;q=0.8",
    "content-type": "application/json",
    "origin": "https://cms.radaa.net",
    "referer": "https://cms.radaa.net/",
    "user-agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
    ),
}

_ALL_TYPES = [
    "FBPAGE_TOPIC", "FBPAGE_COMMENT",
    "FBGROUP_TOPIC", "FBGROUP_COMMENT",
    "FBUSER_TOPIC", "FBUSER_COMMENT",
    "FORUM_TOPIC", "FORUM_COMMENT",
    "NEWS_TOPIC", "NEWS_COMMENT",
    "YOUTUBE_TOPIC", "YOUTUBE_COMMENT",
    "SNS_TOPIC", "SNS_COMMENT",
    "TIKTOK_TOPIC", "TIKTOK_COMMENT",
    "LINKEDIN_TOPIC", "LINKEDIN_COMMENT",
    "ECOMMERCE_TOPIC", "ECOMMERCE_COMMENT",
    "THREADS_TOPIC", "THREADS_COMMENT",
    "REVIEW_TOPIC", "REVIEW_COMMENT",
]

_ALL_SENTIMENTS = ["NONE", "POSITIVE", "NEGATIVE", "NEUTRAL"]

_PAGE_SIZE = 500

# Map type prefix → channel label
_CHANNEL_MAP = {
    "FBPAGE": "Facebook Page",
    "FBGROUP": "Facebook Group",
    "FBUSER": "Facebook User",
    "FORUM": "Forum",
    "NEWS": "News",
    "YOUTUBE": "YouTube",
    "SNS": "SNS",
    "TIKTOK": "TikTok",
    "LINKEDIN": "LinkedIn",
    "ECOMMERCE": "E-commerce",
    "THREADS": "Threads",
    "REVIEW": "Review",
}


def _type_map(raw_type: str) -> str:
    if not raw_type:
        return ""
    
    channel, post_type = raw_type.split("_", 1)

    # Facebook có nhiều loại
    if channel.startswith("FB"):
        fb_map = {
            "FBPAGE": "fbPage",
            "FBGROUP": "fbGroup",
            "FBUSER": "fbUser",
        }
        return fb_map[channel] + post_type.title()

    # Các kênh còn lại
    channel_map = {
        "FORUM": "forum",
        "NEWS": "news",
        "YOUTUBE": "youtube",
        "SNS": "sns",
        "TIKTOK": "tiktok",
        "LINKEDIN": "linkedin",
        "ECOMMERCE": "ecommerce",
        "THREADS": "threads",
        "REVIEW": "review",
    }

    return channel_map[channel] + post_type.title()

_BUZZES_QUERY = """
query buzzes($input: IndexesInput!, $filter: FilterBuzzInput, $filterTotal: FilterBuzzInput) {
  buzzes(input: $input, filter: $filter) {
    status
    message
    total
    skip
    data {
      _id
      _index
      _source {
        type
        publishedDate
        insertedDate
        siteId
        siteName
        url
        title
        content
        description
        likes
        comments
        shares
        interactions
        parentId
        filters
        sentiment { value createdAt createdBy updatedAt updatedBy }
        polarity
        labels { value createdAt createdBy }
        isDeleted
        profile { id name }
        privacy
      }
    }
  }
  trackingTotalBuzzes(input: $input, filter: $filterTotal) {
    status
    message
    total
  }
}
"""


def _ms_to_dt(ms) -> Optional[str]:
    """Convert millisecond timestamp to ISO datetime string (UTC)."""
    if ms is None:
        return None
    try:
        return datetime.fromtimestamp(int(ms) / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(ms)


def _safe_int(val) -> int:
    try:
        return int(val or 0)
    except (ValueError, TypeError):
        return 0


def _map_buzz(item: dict, labels_map: dict[str, str], id_to_name: dict[str, str] = {}) -> dict:
    """Map a raw buzz item ({_id, _index, _source}) to a report-ready flat dict.

    labels_map: { label_id -> label_name } built from project groupTreeLabels.
    """
    src: dict = item.get("_source") or {}
    raw_type: str = src.get("type") or ""

    # _index = "topic{topicId}" → extract topicId → lookup topic name
    index: str = item.get("_index") or ""
    topic_id_from_index = index.removeprefix("topic")
    topic_name = id_to_name.get(topic_id_from_index, "")

    # Derive channel from type prefix (e.g. TIKTOK_COMMENT → TikTok)
    parts = raw_type.rsplit("_", 1)
    channel_key = parts[0] if len(parts) == 2 else raw_type
    channel = _CHANNEL_MAP.get(channel_key, channel_key)

    sentiment_obj = src.get("sentiment") or {}
    labels_list = src.get("labels") or []
    profile = src.get("profile") or {}

    # Resolve label IDs → names via groupTreeLabels lookup
    labels = [
        labels_map.get(lb["value"], lb["value"])
        for lb in labels_list
        if lb and lb.get("value")
    ]

    return {
        "id": item.get("_id"),
        "index": index,
        "topic": topic_name,
        "type": raw_type,
        "channel": channel,
        "content_type": _type_map(raw_type),
        "published_at": _ms_to_dt(src.get("publishedDate")),
        "inserted_at": _ms_to_dt(src.get("insertedDate")),
        "site_id": src.get("siteId"),
        "site_name": src.get("siteName"),
        "url": src.get("url"),
        "title": src.get("title") or "",
        "content": src.get("content") or "",
        "description": src.get("description") or "",
        "author_id": profile.get("id"),
        "author_name": profile.get("name"),
        "likes": _safe_int(src.get("likes")),
        "comments": _safe_int(src.get("comments")),
        "shares": _safe_int(src.get("shares")),
        "interactions": _safe_int(src.get("interactions")),
        "sentiment": sentiment_obj.get("value", "").capitalize() or "",
        "sentiment_by": sentiment_obj.get("createdBy"),
        "polarity": src.get("polarity"),
        "labels": labels,
        "filters": src.get("filters"),
        "parent_id": src.get("parentId"),
        "is_deleted": src.get("isDeleted", False),
        "privacy": src.get("privacy"),
    }


def _fetch_topic_day(
    headers: dict,
    topic_id: str,
    topic_name: str,
    day_from: str,
    day_to: str,
    types: list[str],
    sentiments: list[str],
    extra_filter: Optional[dict],
) -> list[dict]:
    """Fetch tất cả buzzes cho 1 topic trong 1 ngày, dùng trackingTotalBuzzes để tính pages."""
    base_filter: dict = {
        "sorts": {"sortBy": "PUBLISHED_DATE", "sortType": "DESC"},
        "insertedFromDate": day_from,
        "insertedToDate": day_to,
        "types": types,
        "sentiments": sentiments,
        "isDeleted": False,
        "isExactQuery": False,
    }
    if extra_filter:
        base_filter.update(extra_filter)

    def _call(skip: int, track_total: bool) -> dict:
        f  = {**base_filter, "skip": skip, "limit": _PAGE_SIZE}
        ft = {**base_filter, "skip": 0, "limit": _PAGE_SIZE, "trackingTotal": track_total}
        resp = requests.post(
            _CMS_URL,
            json={
                "operationName": "buzzes",
                "variables": {
                    "input": {"indexes": [topic_id]},
                    "filter": f,
                    "filterTotal": ft,
                },
                "query": _BUZZES_QUERY,
            },
            headers=headers,
            timeout=60,
        )
        resp.raise_for_status()
        body = resp.json()
        if body.get("errors"):
            raise ValueError(f"GraphQL errors: {body['errors']}")
        return body.get("data", {})

    # First call: lấy data + total
    first       = _call(skip=0, track_total=True)
    buzzes_resp = first.get("buzzes", {})
    raw_items: list[dict] = buzzes_resp.get("data") or []
    total: int  = first.get("trackingTotalBuzzes", {}).get("total", 0)

    num_pages = math.ceil(total / _PAGE_SIZE) if total else 1
    if total:
        print(f"[fetch_data] {topic_name} {day_from[:10]}: total={total} pages={num_pages} got={len(raw_items)}")

    # Remaining pages
    for page in range(1, num_pages):
        page_data = _call(skip=page, track_total=False)
        items     = page_data.get("buzzes", {}).get("data") or []
        raw_items.extend(items)
        print(f"[fetch_data]   page {page+1}/{num_pages} page={page} got={len(items)} cumulative={len(raw_items)}")
        if not items:
            break

    return raw_items


def fetch_data(
    access_token: str,
    refresh_token: str,
    topic_ids: list[str],
    from_date: str,
    to_date: str,
    topic_id_to_name: Optional[dict[str, str]] = None,
    project_labels: Optional[list[dict]] = None,
    types: Optional[list[str]] = None,
    sentiments: Optional[list[str]] = None,
    extra_filter: Optional[dict] = None,
) -> list[dict]:
    """
    Fetch buzzes từng topic × từng ngày để tránh API limit.
    Mỗi (topic, ngày) dùng trackingTotalBuzzes.total để tính số pages và loop skip.
    """
    labels_map: dict[str, str] = {
        lb["_id"]: lb["name"]
        for lb in (project_labels or [])
        if lb.get("_id") and lb.get("name")
    }
    id_to_name: dict[str, str] = topic_id_to_name or {}
    resolved_types = types or _ALL_TYPES
    resolved_sentiments = sentiments or _ALL_SENTIMENTS

    headers = {
        **_BASE_HEADERS,
        "x-token": f"Bearer {access_token}",
        "x-refresh-token": f"Bearer {refresh_token}",
    }

    # Danh sách từng ngày trong khoảng [from_date, to_date]
    start_dt = datetime.strptime(from_date[:10], "%Y-%m-%d")
    end_dt   = datetime.strptime(to_date[:10],   "%Y-%m-%d")
    dates: list[str] = []
    cur = start_dt
    while cur <= end_dt:
        dates.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)

    all_raw: list[dict] = []

    for topic_id in topic_ids:
        topic_name = id_to_name.get(topic_id, topic_id)
        topic_total = 0
        for date in dates:
            items = _fetch_topic_day(
                headers, topic_id, topic_name,
                f"{date} 00:00:00", f"{date} 23:59:59",
                resolved_types, resolved_sentiments, extra_filter,
            )
            all_raw.extend(items)
            topic_total += len(items)
        print(f"[fetch_data] topic='{topic_name}' subtotal={topic_total}")

    print(f"[fetch_data] grand_total={len(all_raw)}")
    return [_map_buzz(item, labels_map, id_to_name) for item in all_raw]
