"""
llm_client.py
=============
Generate slide text content using DeepInfra via the OpenAI-compatible API.

Env vars:
  DEEPINFRA_API_KEY   API key for DeepInfra
  DEEPINFRA_MODEL     Model ID (default: meta-llama/Llama-3.3-70B-Instruct)

Produces Vietnamese narrative text for:
  slide1_overview    → Slide 1, Rectangle 15  (2-3 câu tổng quan)
  slide1_conclusion  → Slide 1, Rectangle 13  (2 đoạn đúc kết chiến lược)
  slide2_competitive → Slide 2, TextBox 16    (1 đoạn/thương hiệu)
  slide4_analysis    → Slide 4, TextBox 4     (2 đoạn sức khỏe thương hiệu)
"""

from __future__ import annotations
import re
import hashlib
import os
import time

from openai import OpenAI
from dotenv import load_dotenv
load_dotenv()


DEEPINFRA_BASE_URL = "https://api.deepinfra.com/v1/openai"
DEFAULT_MODEL      = "google/gemma-3-12b-it"

# ── In-memory cache: {prompt_hash: (text, expire_ts)} ────────────────────────
_CACHE: dict[str, tuple[str, float]] = {}
_CACHE_TTL = 3600  # 1 hour


def clean_output(text: str) -> str:
    # Xóa label "Đoạn 1:", "Đoạn 2:", "1.", "2." ở đầu đoạn
    text = re.sub(r'(?m)^(Đoạn\s*\d+\s*[:.]?\s*|^\d+[.):]\s*)', '', text)

    # Xóa mọi dòng trống — các đoạn chỉ phân tách bằng đúng 1 ký tự xuống dòng
    text = re.sub(r'\n{2,}', '\n', text)

    return text.strip()


def _format_rules(n_lines: int, max_words: int = 90) -> str:
    """Shared output-format rules block for LLM prompts."""
    return (
        f"## QUY TẮC ĐỊNH DẠNG ĐẦU RA — BẮT BUỘC TUYỆT ĐỐI\n"
        f"- Tổng số dòng trong output phải đúng bằng {n_lines}.\n"
        f"- Độ dài mỗi đoạn từ {max_words - 10} đến {max_words} từ. Nếu vượt quá {max_words}, hãy rút gọn.\n"
        f"- Không được có dòng trống, kể cả ở đầu và cuối output.\n"
        f"- Không được xuống dòng bên trong một đoạn.\n"
        f'- Không có bullet point, markdown, tiêu đề, tiền tố ("Đoạn 1:", "1.", "-",...).\n'
        f"- Không có bất kỳ văn bản nào ngoài {n_lines} đoạn.\n"
        f"\n"
        f"## VÍ DỤ ĐỊNH DẠNG ĐÚNG (minh hoạ cấu trúc, không sao chép nội dung)\n"
        f"<Toàn bộ nội dung đoạn 1 viết liền trên một dòng duy nhất, không xuống dòng.>\n"
        f"<Toàn bộ nội dung đoạn 2 viết liền trên một dòng duy nhất, không xuống dòng.>\n"
        f"\n"
        f"## VÍ DỤ ĐỊNH DẠNG SAI — TUYỆT ĐỐI KHÔNG LÀM THEO\n"
        f"Đoạn 1: Nội dung...\n"
        f"\n"
        f"Nội dung tiếp theo...\n"
        f"\n"
        f"Đoạn 2: Nội dung..."
    )


def _cache_get(key: str) -> str | None:
    entry = _CACHE.get(key)
    if entry and time.time() < entry[1]:
        return entry[0]
    _CACHE.pop(key, None)
    return None


def _cache_set(key: str, value: str) -> None:
    _CACHE[key] = (value, time.time() + _CACHE_TTL)


def _cache_key(prompt: str) -> str:
    return hashlib.sha256(prompt.encode()).hexdigest()


# ── Low-level caller ──────────────────────────────────────────────────────────

def call_llm(api_key: str, prompt: str, max_tokens: int = 800) -> str:
    key = _cache_key(f"{max_tokens}|{prompt}")
    cached = _cache_get(key)
    if cached is not None:
        print("⚡ LLM cache hit")
        return cached
    model = os.getenv("DEEPINFRA_MODEL", DEFAULT_MODEL)
    client = OpenAI(api_key=api_key, base_url=DEEPINFRA_BASE_URL)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=max_tokens,
        temperature=0.7,
        
    )
    result = response.choices[0].message.content.strip()
    result = clean_output(result)
    print(result)
    _cache_set(key, result)
    return result


# ── Prompt builders ───────────────────────────────────────────────────────────

def _build_topic_detail_block(topic_health_raw: dict, brand_total: int) -> str:
    """Format top-6 topics as bullet lines with % share and sentiment."""
    lines = []
    denom = brand_total or 1
    for topic, v in list(topic_health_raw.items())[:6]:
        grand_pct = round(v["total"] / denom * 100, 1)
        lines.append(
            f"{topic}: {grand_pct}% tổng thảo luận"
        )
    return " | ".join(lines) or "  (chưa có dữ liệu phân loại chủ đề)"


_SENTIMENT_VI = {"positive": "Tích cực", "negative": "Tiêu cực", "neutral": "Trung lập"}

def _truncate(posts, max_words: int = 200, max_comments: int = 0) -> str:
    if not posts:
        return "N/A"
    if isinstance(posts, str):
        return posts
    texts = []
    for post in posts:
        if not post:
            continue
        seen: set[str] = set()
        parts: list[str] = []
        for col in ["Title", "Content", "Description"]:
            val = (post.get(col) or "").strip() if isinstance(post, dict) else str(post)
            if val and val not in seen:
                seen.add(val)
                parts.append(val)
        full_text = " ".join(" | ".join(parts).split()[:max_words])
        sentiment = post.get("Sentiment", "") if isinstance(post, dict) else ""
        label = _SENTIMENT_VI.get(sentiment.lower(), "")
        entry = f"[{label}] {full_text}" if label else full_text
        if max_comments and isinstance(post, dict):
            comments = (post.get("Comment") or [])[:max_comments]
            if comments:
                comment_str = " | ".join(" ".join(c.split()[:20]) for c in comments)
                entry += f"\n  Bình luận: {comment_str}"
        texts.append(entry)
    return "\n".join(texts) or "N/A"

def _build_comp_block(all_brands: list[str], brand_metrics: dict, top_pos: int) -> str:
    """Format one block per brand: rank + prev rank + top topic + top post with sentiment label."""
    ranked_curr = sorted(all_brands, key=lambda b: brand_metrics.get(b, {}).get("current", {}).get("total", 0), reverse=True)
    rank_of = {b: i + 1 for i, b in enumerate(ranked_curr)}

    ranked_prev = sorted(all_brands, key=lambda b: brand_metrics.get(b, {}).get("previous", {}).get("total", 0), reverse=True)
    prev_rank_of = {b: i + 1 for i, b in enumerate(ranked_prev)}

    blocks = []
    for brand in all_brands:
        bm        = brand_metrics.get(brand, {})
        topic_h   = bm.get("topic_health", {})
        top_topic = next(iter(topic_h), "chưa phân loại")
        posts     = bm.get("top_posts", [])
        top_post  = _truncate(posts[:top_pos], 1000) if posts else "N/A"
        blocks.append(
            f"- {brand} (Kỳ này: #{rank_of[brand]}, Kỳ trước: #{prev_rank_of[brand]})\n"
            f"  Chủ đề nổi bật: {top_topic}\n"
            f"  Nội dung nổi bật: {top_post}"
        )
    return "\n".join(blocks)


# ── Main entry point ──────────────────────────────────────────────────────────

_NO_DATA_TEXTS = {
    "slide1_overview":    "Không ghi nhận thảo luận nào trong kỳ báo cáo này.",
    "slide1_conclusion":  "Không ghi nhận thảo luận nào trong kỳ báo cáo này. Không có dữ liệu để phân tích điểm mạnh, cơ hội hay rủi ro.",
    "slide2_competitive": "Không ghi nhận thảo luận nào từ các thương hiệu trong kỳ báo cáo này.",
    "slide4_analysis":    "Không ghi nhận thảo luận nào trong kỳ báo cáo này. Không có dữ liệu để phân tích sức khỏe thương hiệu.",
}


def generate_slide_texts(data: dict, api_key: str) -> dict:
    """
    Call DeepInfra to generate narrative text for slides 1, 2, 4.

    Returns dict with keys:
      slide1_overview, slide1_conclusion, slide2_competitive, slide4_analysis
    Values are text strings, or None if the call failed.
    """
    ctx         = data["llm_context"]
    brand       = ctx["brand"]
    period      = ctx["period"]
    total       = ctx["total_current"]

    if not total:
        print("ℹ️  No discussion data — skipping LLM, using no-data fallback texts.")
        return dict(_NO_DATA_TEXTS)
    pos_pct     = ctx["pos_pct"]
    neg_pct     = ctx["neg_pct"]
    neu_pct     = round(100 - pos_pct - neg_pct, 1)
    nsr         = ctx["nsr"]
    top_channel = ctx["top_channel"]
    top_ch_pct  = ctx["top_channel_pct"]
    pos_verbatims       = ctx.get("pos_verbatims", None)
    neg_verbatims       = ctx.get("neg_verbatims", None)
    top_sources     = ctx.get("top_sources", [])
    topic_line      = ctx.get("topic_line", "")
    data_trend_line = ctx.get("data_trend_line", "")



    pos_content = _truncate(pos_verbatims, max_comments=3)
    neg_content = _truncate(neg_verbatims, max_comments=3)

    all_brands       = data["all_brands"]
    brand_metrics    = data["brand_metrics"]
    topic_health_raw = data["charts"].get("topic_health", {}).get("raw", {})

    topic_detail_block = _build_topic_detail_block(topic_health_raw, total)
    comp_block         = _build_comp_block(all_brands, brand_metrics, 1)

    # ── Slide 1: Conclusion (Rectangle 13) ───────────────────────────────────
    prompt_conclusion = f"""
Bạn là chuyên gia viết báo cáo thương hiệu.
Dựa vào dữ liệu được cung cấp hãy viết nội dung đúc kết tiếng Việt cho báo cáo thương hiệu Masan Group.

## NHIỆM VỤ
Viết đúng 2 đoạn văn xuôi liên tục, không ngắt dòng trong mỗi đoạn. Hai đoạn chỉ được phân tách bằng đúng một ký tự xuống dòng duy nhất.

## NỘI DUNG TỪNG ĐOẠN
Đoạn 1: Điểm mạnh hoặc cơ hội cần phát huy, trình bày ngắn gọn dẫn chứng cụ thể từ dữ liệu), đề xuất hành động để duy trì.
Đoạn 2: Các vấn đề hoặc rủi ro cần xử lý, trình bày ngắn gọn nguyên nhân, đưa ra khuyến nghị hành động rõ ràng.

{_format_rules(2, max_words=65)}

### DỮ LIỆU
- Nội dung tích cực nổi bật: {pos_content}
- Nội dung tiêu cực nổi bật: {neg_content}

"""

    # ── Slide 2: Competitive landscape (TextBox 16) ──────────────────────────
    _curr_top1 = max(all_brands, key=lambda b: brand_metrics.get(b, {}).get("current", {}).get("total", 0))
    _prev_top1 = max(all_brands, key=lambda b: brand_metrics.get(b, {}).get("previous", {}).get("total", 0))
    if _curr_top1 == _prev_top1:
        _top1_instruction = f'Với {_curr_top1} (đứng #1 kỳ này và cũng #1 kỳ trước): bắt đầu đoạn bằng "{_curr_top1} tiếp tục dẫn đầu..."'
    else:
        _top1_instruction = f'Với {_curr_top1} (đứng #1 kỳ này, kỳ trước không phải #1): bắt đầu đoạn bằng "{_curr_top1} vươn lên dẫn đầu..."'

    prompt_slide2 = f"""
Bạn là chuyên gia viết báo cáo thương hiệu.
Dựa vào dữ liệu được cung cấp hãy viết nội dung đúc kết tiếng Việt cho báo cáo thương hiệu Masan Group.

## NHIỆM VỤ
Viết đúng các đoạn văn xuôi liên tục, không ngắt dòng trong mỗi đoạn. Các đoạn chỉ được phân tách bằng đúng một ký tự xuống dòng duy nhất.

## NỘI DUNG TỪNG ĐOẠN
Mỗi thương hiệu trình bày 1 đoạn, mở đầu bằng tên thương hiệu, nội dung nổi bật suy ra từ chủ đề và bài đăng nổi bật. Không đề cập thứ hạng hay vị thế của các thương hiệu không phải dẫn đầu.
{_top1_instruction}

{_format_rules(len(all_brands), max_words=25)}

### DỮ LIỆU
{comp_block}
"""

    # ── Slide 4: Brand health analysis (TextBox 4) ───────────────────────────

    prompt_slide4 = f"""Bạn là chuyên gia viết báo cáo thương hiệu.
Dựa vào dữ liệu được cung cấp hãy viết nội dung tổng kết tiếng Việt cho báo cáo thương hiệu Masan Group.

## NHIỆM VỤ
Viết đúng 2 đoạn văn xuôi liên tục, không ngắt dòng trong mỗi đoạn. Hai đoạn chỉ được phân tách bằng đúng một ký tự xuống dòng duy nhất.

## NỘI DUNG TỪNG ĐOẠN
Đoạn 1: Mô tả bức tranh thảo luận tổng thể. Phân tích các chủ đề chiếm tỷ trọng cao nhất. Đề cập sắc thái áp đảo và sức khoẻ thương hiệu. Đề cập nguồn thảo luận lớn nhất.
Đoạn 2: Phân tích sắc thái thảo luận. Liên hệ với bài đăng tích cực và tiêu cực nổi bật. Phân tích xu hướng thảo luận theo thời gian. Kết thúc bằng một nhận định ngắn gọn dựa trên dữ liệu.

{_format_rules(2, max_words=65)}

### DỮ LIỆU:
- Sức khoẻ thương hiệu (NSR = (Tích cực - Tiêu cực) / (Tích cực + Tiêu cực)): {nsr}%
- Phân bổ sắc thái: Tích cực: {pos_pct}%, Trung lập: {neu_pct}%, Tiêu cực: {neg_pct}%
- Nguồn lớn nhất: {top_channel} ({top_ch_pct}%)
- Phân bổ chủ đề:
{topic_detail_block}
- Nội dung tích cực nổi bật: {pos_content}
- Nội dung tiêu cực nổi bật: {neg_content}
-  Xu hướng thảo luận {data_trend_line}

    """

    # ── Call LLM for each prompt ──────────────────────────────────────────────
    prompts = {
        "slide1_conclusion":  (prompt_conclusion,  250),
        "slide2_competitive": (prompt_slide2,      280),
        "slide4_analysis":    (prompt_slide4,      250),
    }

    texts: dict[str, str | None] = {}
    for key, (prompt, max_toks) in prompts.items():
        try:
            texts[key] = call_llm(api_key, prompt, max_toks)
            print(f"✅ LLM generated: {key}")
        except Exception as exc:
            print(f"⚠️  LLM failed for {key}: {exc}")
            texts[key] = None

    return texts
