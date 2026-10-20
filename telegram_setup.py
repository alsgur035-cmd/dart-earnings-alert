"""텔레그램 봇 연결 도우미 (PC에서 한 번만 실행).

1) 봇 토큰을 붙여넣으면 2) 내 채팅 ID(TELEGRAM_CHAT_ID)를 찾아 화면에 보여주고 3) 시험 메시지를 보낸다.
실행 전에 텔레그램 앱에서 만든 봇에게 아무 메시지나 한 번 보내 둘 것(그래야 채팅 ID가 잡힌다).
토큰은 파일에 저장하지 않는다 — GitHub Secrets에만 넣는다.
"""
from __future__ import annotations

import json
import sys
from urllib.parse import urlencode
from urllib.request import Request, urlopen


def call(token: str, method: str, params: dict | None = None) -> dict:
    data = urlencode(params).encode() if params else None
    with urlopen(Request(f"https://api.telegram.org/bot{token}/{method}", data=data), timeout=20) as resp:
        return json.loads(resp.read().decode())


def main() -> int:
    token = input("BotFather가 준 봇 토큰을 붙여넣고 Enter: ").strip()
    me = call(token, "getMe")
    if not me.get("ok"):
        print("토큰이 올바르지 않습니다:", me)
        return 1
    print(f"봇 확인: @{me['result']['username']}")

    updates = call(token, "getUpdates").get("result", [])
    chats = {}
    for u in updates:
        chat = (u.get("message") or u.get("channel_post") or {}).get("chat")
        if chat:
            chats[chat["id"]] = chat.get("first_name") or chat.get("title") or ""
    if not chats:
        print("받은 메시지가 없습니다. 텔레그램에서 봇에게 아무 말이나 보낸 뒤 다시 실행하세요.")
        return 1

    for chat_id, name in chats.items():
        print(f"\nTELEGRAM_CHAT_ID = {chat_id}   ({name})")
        call(token, "sendMessage", {"chat_id": chat_id, "text": "✅ DART 실적 알림 연결 시험 메시지입니다."})
    print("\n폰으로 시험 메시지가 왔으면 위 숫자를 GitHub Secrets의 TELEGRAM_CHAT_ID에 넣으세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
