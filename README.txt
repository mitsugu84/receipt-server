Receipt Server v1

Render用の領収書解析サーバーです。

構成:
PCアプリ
↓
Renderサーバー
↓
OpenAI API
↓
JSON返却

APIキーはRenderの環境変数に保存するため、PCアプリ側にOpenAI APIキーを置かなくて済みます。

----------------------------------------
ファイル構成
----------------------------------------

main.py
requirements.txt
render.yaml
README.txt

----------------------------------------
Renderデプロイ手順
----------------------------------------

1. GitHubにこのフォルダをアップロード

2. Renderで New → Web Service

3. GitHubリポジトリを選択

4. 設定

Build Command:
pip install -r requirements.txt

Start Command:
uvicorn main:app --host 0.0.0.0 --port $PORT

5. Environment Variables に追加

OPENAI_API_KEY = 自分のOpenAI APIキー
OPENAI_MODEL = gpt-4.1-mini

6. Deploy

----------------------------------------
動作確認
----------------------------------------

ブラウザで以下を開く。

https://あなたのRenderURL/health

例:
{
  "status": "ok",
  "model": "gpt-4.1-mini",
  "openai_key": "set"
}

----------------------------------------
画像解析API
----------------------------------------

POST /analyze

form-data:
file = 領収書画像

返却JSON例:
{
  "date": "2026/06/04",
  "vendor": "キヌヤ",
  "store": "キヌヤ高陽店",
  "amount": 2633,
  "category": "食費",
  "memo": "食品購入"
}

----------------------------------------
注意
----------------------------------------

Renderの無料プランは、しばらくアクセスがないとスリープします。
初回アクセス時だけ起動に時間がかかることがあります。
