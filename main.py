# -*- coding: utf-8 -*-
"""
Receipt Server v1

Render / FastAPI 用サーバー

機能:
- GET /health
- POST /analyze
- 領収書画像を受け取る
- OpenAI Visionで解析
- 日付 / 取引先 / 店名 / 金額 / 区分 / メモ をJSONで返す
"""

import base64
import json
import os
import traceback
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from openai import OpenAI
from pydantic import BaseModel


APP_NAME = "Receipt Server v1"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini").strip() or "gpt-4.1-mini"

DEFAULT_CATEGORIES = [
    "材料費",
    "工具費",
    "燃料費",
    "高速料金",
    "駐車料金",
    "宿泊費",
    "食費",
    "消耗品費",
    "通信費",
    "雑費",
    "その他",
]


app = FastAPI(
    title=APP_NAME,
    version="1.0.0",
    description="Receipt image analysis server using OpenAI Vision.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ReceiptResult(BaseModel):
    date: str
    vendor: str
    store: str
    amount: int
    category: str
    memo: str


class ErrorResult(BaseModel):
    error: str
    detail: str | None = None


def get_openai_client() -> OpenAI:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not set.")
    return OpenAI(api_key=OPENAI_API_KEY)


def guess_mime_type(filename: str, content_type: str | None) -> str:
    if content_type and content_type.startswith("image/"):
        return content_type

    lower = filename.lower()

    if lower.endswith(".png"):
        return "image/png"
    if lower.endswith(".webp"):
        return "image/webp"
    if lower.endswith(".bmp"):
        return "image/bmp"
    if lower.endswith(".heic") or lower.endswith(".heif"):
        return "image/heic"
    if lower.endswith(".jpg") or lower.endswith(".jpeg"):
        return "image/jpeg"

    return "image/jpeg"


def bytes_to_data_url(image_bytes: bytes, mime_type: str) -> str:
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime_type};base64,{b64}"


def parse_json_from_text(text: str) -> dict[str, Any]:
    text = text.strip()

    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].strip()

    start = text.find("{")
    end = text.rfind("}")

    if start >= 0 and end >= 0:
        text = text[start : end + 1]

    return json.loads(text)


def safe_int(value: Any) -> int:
    try:
        if value is None:
            return 0

        if isinstance(value, str):
            value = value.replace(",", "")
            value = value.replace("円", "")
            value = value.replace("¥", "")
            value = value.strip()

        return int(float(value))

    except Exception:
        return 0


def normalize_date(value: Any) -> str:
    value = str(value or "").strip()

    if not value:
        return ""

    value = value.replace("-", "/")
    value = value.replace(".", "/")
    value = value.replace("年", "/")
    value = value.replace("月", "/")
    value = value.replace("日", "")

    parts = [p.strip() for p in value.split("/") if p.strip()]

    if len(parts) != 3:
        return value

    try:
        y, m, d = parts

        if len(y) == 2:
            y_int = int(y)
            y = f"20{y}" if y_int < 80 else f"19{y}"

        return f"{int(y):04d}/{int(m):02d}/{int(d):02d}"

    except Exception:
        return value


def build_prompt() -> str:
    categories_text = "、".join(DEFAULT_CATEGORIES)

    return f"""
あなたは日本の領収書・レシート入力アシスタントです。
画像から、税理士・会計士に渡すExcel用の情報を抽出してください。

必ずJSONのみで返してください。説明文は不要です。

出力形式:
{{
  "date": "YYYY/MM/DD または 空欄",
  "vendor": "取引先名。例: キヌヤ、ENEOS、コーナン、ローソン。不明なら空欄",
  "store": "店名・支店名・会社名。例: キヌヤ高陽店。不明なら空欄",
  "amount": 数値のみ。税込合計金額・領収金額。不明なら0,
  "category": "次の中から最も近いもの: {categories_text}",
  "memo": "短いメモ。例: 食品購入、ガソリン給油、工具購入、材料購入。不明なら空欄"
}}

分類ルール:
- スーパー、食品店、飲食店、コンビニで食品・弁当・飲料中心 → 食費
- ガソリンスタンド、軽油、ガソリン、給油 → 燃料費
- ホームセンター、資材、部材、ネジ、電材、配管材 → 材料費
- 工具、ドリル、刃、レンチ、電動工具 → 工具費
- コピー用紙、文具、日用品、掃除用品 → 消耗品費
- 高速道路、ETC、有料道路 → 高速料金
- 駐車場、パーキング → 駐車料金
- ホテル、旅館、宿泊 → 宿泊費
- 携帯、通信、ネット → 通信費
- 判断できないもの → その他

注意:
- 金額は小計ではなく、合計・税込合計・領収金額・支払金額を最優先してください。
- 税額、外税、内税、お釣り、ポイント、バーコード番号、登録番号はamountにしないでください。
- amountにカンマや円マークは含めないでください。
- 日付が年なしの場合は、画像内やファイル情報から推定せず空欄で構いません。
- vendorは短い取引先名、storeはできるだけ正式な店名にしてください。
""".strip()


def analyze_receipt_image(
    image_bytes: bytes,
    filename: str,
    content_type: str | None,
) -> ReceiptResult:
    client = get_openai_client()

    mime_type = guess_mime_type(filename, content_type)
    data_url = bytes_to_data_url(image_bytes, mime_type)
    prompt = build_prompt()

    response = client.responses.create(
        model=OPENAI_MODEL,
        input=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": prompt,
                    },
                    {
                        "type": "input_image",
                        "image_url": data_url,
                    },
                ],
            }
        ],
    )

    result_text = response.output_text
    data = parse_json_from_text(result_text)

    return ReceiptResult(
        date=normalize_date(data.get("date", "")),
        vendor=str(data.get("vendor", "") or ""),
        store=str(data.get("store", "") or ""),
        amount=safe_int(data.get("amount", 0)),
        category=str(data.get("category", "") or ""),
        memo=str(data.get("memo", "") or ""),
    )


@app.get("/")
def root() -> dict[str, str]:
    return {
        "name": APP_NAME,
        "status": "ok",
        "usage": "POST /analyze with form-data file=<image>",
    }


@app.get("/health")
def health() -> dict[str, str]:
    return {
        "status": "ok",
        "model": OPENAI_MODEL,
        "openai_key": "set" if bool(OPENAI_API_KEY) else "missing",
    }


@app.post(
    "/analyze",
    response_model=ReceiptResult,
    responses={500: {"model": ErrorResult}},
)
async def analyze(file: UploadFile = File(...)) -> ReceiptResult:
    try:
        if not file.filename:
            raise HTTPException(status_code=400, detail="filename is empty")

        image_bytes = await file.read()

        if not image_bytes:
            raise HTTPException(status_code=400, detail="file is empty")

        return analyze_receipt_image(
            image_bytes=image_bytes,
            filename=file.filename,
            content_type=file.content_type,
        )

    except HTTPException:
        raise

    except Exception as e:
        detail = traceback.format_exc()
        raise HTTPException(
            status_code=500,
            detail={
                "error": str(e),
                "traceback": detail[-3000:],
            },
        )
