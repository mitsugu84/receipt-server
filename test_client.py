# -*- coding: utf-8 -*-
"""
ローカルテスト用クライアント

使い方:
python test_client.py https://あなたのRenderURL/analyze receipt.jpg
"""

import sys
import requests


def main():
    if len(sys.argv) < 3:
        print("Usage: python test_client.py <server_analyze_url> <image_path>")
        print("Example: python test_client.py https://example.onrender.com/analyze receipt.jpg")
        return

    url = sys.argv[1]
    image_path = sys.argv[2]

    with open(image_path, "rb") as f:
        files = {
            "file": (image_path, f, "image/jpeg"),
        }
        res = requests.post(url, files=files, timeout=120)

    print("status:", res.status_code)
    print(res.text)


if __name__ == "__main__":
    main()
