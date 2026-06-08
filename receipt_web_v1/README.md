# 領収書まとめ Web版 v1

## できること

- スマホ・PCブラウザから領収書画像をアップロード
- AIで日付・店名・金額・区分を解析
- Excelを自動生成
- Excelをダウンロード

## ローカル起動

```bash
pip install -r requirements.txt
```

OpenAI APIキーを設定します。

Windows PowerShell:

```powershell
$env:OPENAI_API_KEY="sk-xxxx"
python app.py
```

ブラウザで開きます。

```text
http://localhost:5000
```

## Renderデプロイ

Renderで新しいWeb Serviceを作り、このフォルダをGitHubに上げて接続します。

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
gunicorn app:app
```

Environment Variables:

```text
OPENAI_API_KEY=sk-xxxx
OPENAI_MODEL=gpt-4o-mini
FLASK_SECRET_KEY=好きな長い文字列
```

## 次に追加する機能

1. メール送信
2. ユーザー登録
3. 月20枚制限
4. 月末自動Excel送信
5. Stripe課金
