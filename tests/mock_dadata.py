# -*- coding: utf-8 -*-
"""
Мини-эмулятор DaData для тестов: отвечает на POST /suggestions/api/4_1/rs/findById/party.

Поведение задаётся по ИНН (см. BEHAVIORS) и позволяет проверить:
успешный ответ, отсутствие организации, повторы при 429/500, ошибку авторизации,
исчерпание квоты, невалидный JSON и «обычную» HTTP-ошибку.

Дополнительно сервер собирает статистику обращений (GET /__stats), по которой
тесты проверяют соблюдение лимитов: времена запросов и запрошенные ИНН.
"""

from __future__ import annotations

import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Optional


def _suggestion(
    inn: str,
    name: str,
    okved_main: str,
    okveds: Any,
    ogrn: Optional[str] = None,
    kpp: Optional[str] = None,
    party_type: str = "LEGAL",
) -> dict[str, Any]:
    """
    Формирует ответ в формате DaData suggestions.

    Как и настоящий API: ОГРН и ОГРНИП приходят в одном поле data.ogrn
    (13 или 15 цифр), отдельного data.ogrnip нет. Для ИП задавайте ogrn
    из 15 цифр и party_type="INDIVIDUAL", для юрлица — 13 цифр и "LEGAL".
    """
    ogrn_value = ogrn if ogrn is not None else "1" + inn[::-1][:12].ljust(12, "0")
    data: dict[str, Any] = {
        "inn": inn,
        "ogrn": ogrn_value,
        "type": party_type,
        "okved": okved_main,
        "okveds": okveds,
        "name": {"short_with_opf": name},
    }
    if kpp is not None:
        data["kpp"] = kpp
    return {
        "suggestions": [
            {
                "value": name,
                "unrestricted_value": name,
                "data": data,
            }
        ]
    }


# Запрос (ИНН, ОГРН или ОГРНИП) -> список реакций (по порядку обращений);
# последняя используется и далее. Ключи — те же значения, что уходят в {"query": ...}.
BEHAVIORS: dict[str, list[dict[str, Any]]] = {
    "7707083893": [{"kind": "json", "body": _suggestion("7707083893", 'ООО "Ромашка"', "56.10",
                                                        [{"code": "46.90", "main": False},
                                                         {"code": "62.01", "main": False}],
                                                        ogrn="1027700000000",
                                                        kpp="770701001")}],
    "500100732259": [{"kind": "json", "body": _suggestion("500100732259", 'ИП Кузнецов А.А.', "46.73",
                                                          [{"code": "43.21"}],
                                                          ogrn="326547600194488",
                                                          party_type="INDIVIDUAL")}],
    "7830002293": [{"kind": "json", "body": _suggestion("7830002293", 'ООО "Лютик"', "62.01",
                                                        ["71.12", "46.90"])}],
    # ОГРН/ОГРНИП-ключи: по ним тоже ищут (findById/party принимает оба формата)
    "1027700000000": [{"kind": "json", "body": _suggestion("7707083893", 'ООО "Ромашка"', "56.10",
                                                           [{"code": "46.90", "main": False}],
                                                           ogrn="1027700000000",
                                                           kpp="770701001")}],
    "326547600194488": [{"kind": "json", "body": _suggestion("500100732259", 'ИП Кузнецов А.А.', "46.73",
                                                             [{"code": "43.21"}],
                                                             ogrn="326547600194488",
                                                             party_type="INDIVIDUAL")}],
    "1234567890": [{"kind": "json", "body": {"suggestions": []}}],
    # ИНН есть в списке, но в реестре не найден, а ОГРН из той же строки — находится
    "1027700009999": [{"kind": "json", "body": _suggestion("7707083893", 'ООО "Ромашка"', "56.10",
                                                           [{"code": "46.90"}],
                                                           ogrn="1027700009999")}],
    "9999999999": [{"kind": "status", "code": 500, "text": "internal error"},
                   {"kind": "status", "code": 500, "text": "internal error"},
                   {"kind": "json", "body": _suggestion("9999999999", 'ООО "Пятый"', "56.20", [])}],
    "8888888888": [{"kind": "status", "code": 429, "text": "too many requests",
                    "headers": {"Retry-After": "1"}},
                   {"kind": "json", "body": _suggestion("8888888888", 'ООО "Шестой"', "47.11", [])}],
    "4444444444": [{"kind": "json", "body": _suggestion("4444444444", 'ООО "Тринадцатый"', "46.01",
                                                        None)}],
    "6666666666": [{"kind": "raw", "code": 200, "text": "this is not json"}],
    "5555555555": [{"kind": "status", "code": 400, "text": "bad request"}],
    "7777777777": [{"kind": "status", "code": 403,
                    "text": "Достигнут лимит на количество запросов в сутки"}],
    "3333333333": [{"kind": "status", "code": 403, "text": "Ваш ключ не имеет доступа"}],
}

DEFAULT_BEHAVIOR: list[dict[str, Any]] = [
    {"kind": "json", "body": _suggestion("0000000000", 'ООО "Тестовая"', "62.01", [])}
]


class Stats:
    """Собирает обращения к мок-серверу (значение query: ИНН, ОГРН или ОГРНИП)."""

    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.times: list[float] = []
        self.queries: list[str] = []
        self.statuses: list[int] = []

    def add(self, query: str, status: int) -> None:
        with self.lock:
            self.times.append(time.monotonic())
            self.queries.append(query)
            self.statuses.append(status)

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return {
                "total": len(self.times),
                "times": list(self.times),
                "queries": list(self.queries),
                "inns": list(self.queries),  # прежнее имя ключа, оставлено для совместимости
                "statuses": list(self.statuses),
            }


class _Handler(BaseHTTPRequestHandler):
    stats: Stats = Stats()
    attempts: dict[str, int] = {}
    behavior: dict[str, list[dict[str, Any]]] = BEHAVIORS

    def log_message(self, *args: Any) -> None:  # тишина в тестах
        return

    def _send(self, code: int, body: str, headers: Optional[dict[str, str]] = None,
              content_type: str = "application/json") -> None:
        payload = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:  # noqa: N802
        if self.path.startswith("/__stats"):
            self._send(200, json.dumps(self.stats.snapshot(), ensure_ascii=False))
            return
        self._send(404, json.dumps({"error": "not found"}))

    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        try:
            inn = str(json.loads(raw).get("query") or "")
        except ValueError:
            inn = ""

        behavior_list = self.behavior.get(inn, DEFAULT_BEHAVIOR)
        index = self.attempts.get(inn, 0)
        self.attempts[inn] = index + 1
        action = behavior_list[min(index, len(behavior_list) - 1)]

        status = int(action.get("code", 200))
        self.stats.add(inn, status)

        if action["kind"] == "json":
            self._send(status, json.dumps(action["body"], ensure_ascii=False))
        elif action["kind"] == "raw":
            self._send(status, str(action.get("text", "")), content_type="text/plain")
        else:  # status
            self._send(status, json.dumps({"error": action.get("text", "")}, ensure_ascii=False),
                       headers=action.get("headers"))


class MockDaData:
    """Запускает мок-сервер на свободном порту и отдаёт его URL."""

    def __init__(self, behaviors: Optional[dict[str, list[dict[str, Any]]]] = None) -> None:
        self.stats = Stats()
        handler = type("_BoundHandler", (_Handler,), {
            "stats": self.stats,
            "attempts": {},
            "behavior": behaviors if behaviors is not None else BEHAVIORS,
        })
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    @property
    def url(self) -> str:
        host, port = self.server.server_address[:2]
        return f"http://{host}:{port}/suggestions/api/4_1/rs/findById/party"

    def __enter__(self) -> "MockDaData":
        self.thread.start()
        return self

    def __exit__(self, *exc: Any) -> None:
        self.server.shutdown()
        self.server.server_close()

    def snapshot(self) -> dict[str, Any]:
        return self.stats.snapshot()


if __name__ == "__main__":  # ручной запуск: python3 tests/mock_dadata.py
    with MockDaData() as mock:
        print("Мок DaData слушает:", mock.url)
        print("Статистика: GET http://127.0.0.1:%d/__stats" % mock.server.server_address[1])
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
