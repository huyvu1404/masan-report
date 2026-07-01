"""
CMS authentication and project info client (GraphQL).

Required .env vars:
    CMS_URL        GraphQL endpoint, e.g. https://cms-gateway.radaa.net/kompaql
    CMS_USERNAME   Login username
    CMS_PASSWORD   Login password
"""

from __future__ import annotations

import os
from typing import Optional

import requests
from dotenv import load_dotenv

load_dotenv()

_CMS_URL = os.getenv("CMS_URL", "").rstrip("/")
_USERNAME = os.getenv("CMS_USERNAME", "")
_PASSWORD = os.getenv("CMS_PASSWORD", "")

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

_LOGIN_MUTATION = """
mutation login($input: LoginInput!) {
  login(input: $input) {
    status
    message
    accessToken
    refreshToken
    data {
      _id
      username
      firstName
      lastName
      email
    }
  }
}
"""

_PROJECT_QUERY = """
query project($_id: ID!) {
  project(_id: $_id) {
    status
    message
    data {
      _id
      domain
      name
      displayName
      status
      labels {
        _id
        name
        path
      }
      groupTreeLabels {
        _id
        name
        path
      }
      filters {
        _id
        name
        type
      }
      topics {
        _id
        name
      }
    }
  }
}
"""


def _gql(payload: dict, access_token: Optional[str] = None, refresh_token: Optional[str] = None) -> dict:
    headers = dict(_BASE_HEADERS)
    if access_token:
        headers["x-token"] = f"Bearer {access_token}"
    if refresh_token:
        headers["x-refresh-token"] = f"Bearer {refresh_token}"
    resp = requests.post(_CMS_URL, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    body = resp.json()
    if body.get("errors"):
        raise ValueError(f"GraphQL errors: {body['errors']}")
    return body.get("data", {})


def login() -> tuple[str, str]:
    """
    Authenticate and return (access_token, refresh_token).
    """
    data = _gql({
        "operationName": "login",
        "variables": {"input": {"username": _USERNAME, "password": _PASSWORD}},
        "query": _LOGIN_MUTATION,
    })
    login_data = data.get("login", {})
    access_token = login_data.get("accessToken")
    refresh_token = login_data.get("refreshToken", "")
    if not access_token:
        raise ValueError(f"No accessToken in login response: {data}")
    return access_token, refresh_token


def get_project_info(
    project_id: str,
    access_token: str,
    refresh_token: str = "",
    selected_topics: Optional[list[str]] = None,
    selected_topic_ids: Optional[list[str]] = None,
    meta_report: Optional[dict] = None,
) -> dict:
    """
    Fetch project info and return a structured dict:
        { projectId, projectName, labels, filters, topics }

    topics filtered by selected_topic_ids (priority) or selected_topics.
    Each topic entry includes fileName when meta_report={'type':..., 'period_str':...}.
    """
    data = _gql(
        {
            "operationName": "project",
            "variables": {"_id": project_id},
            "query": _PROJECT_QUERY,
        },
        access_token=access_token,
        refresh_token=refresh_token,
    )

    project_data: dict = data.get("project", {}).get("data") or {}

    project_id_val = project_data.get("_id", project_id)
    project_name = project_data.get("name", "")
    group_tree_labels: list = project_data.get("groupTreeLabels", [])
    raw_filters: list = project_data.get("filters", [])
    all_topics: list = project_data.get("topics", [])

    # groupTreeLabels is a nested list [[{_id, name, path}, ...], [...]]
    # mirror the JS: groupTreeLabels.flat().filter(x => x?.path)
    def _flat(lst):
        for el in lst:
            if isinstance(el, list):
                yield from _flat(el)
            else:
                yield el

    labels = [
        {"_id": item["_id"], "name": item["name"], "path": item["path"]}
        for item in _flat(group_tree_labels)
        if item.get("path")
    ]

    filters = [{"_id": f["_id"], "name": f["name"]} for f in raw_filters]

    topics_data = all_topics
    if selected_topic_ids:
        topics_data = [t for t in all_topics if t["_id"] in selected_topic_ids]
    elif selected_topics:
        topics_data = [t for t in all_topics if t["name"] in selected_topics]

    if meta_report:
        report_type = meta_report.get("type", "")
        period_str = meta_report.get("period_str", "")
        topics = [
            {
                "topic": t["name"],
                "topicId": t["_id"],
                "fileName": f"{t['_id']}-{report_type}-{period_str}.xlsx",
            }
            for t in topics_data
        ]
    else:
        topics = [{"topic": t["name"], "topicId": t["_id"]} for t in topics_data]

    return {
        "projectId": project_id_val,
        "projectName": project_name,
        "labels": labels,
        "filters": filters,
        "topics": topics,
    }
