#!/usr/bin/env python3
"""
Simple checkout API endpoint for Capstone AI articles.

Features:
- Accepts POST JSON at /api/capstone-checkout
- Handles CORS (including Origin: null from file:// pages)
- Logs every event to CSV
- Maintains current checkout state in a second CSV

Run:
    python Python/capstone_checkout_endpoint.py

Then set in your HTML:
    CHECKOUT_UPLOAD_URL = 'http://127.0.0.1:8765/api/capstone-checkout'
"""

from __future__ import annotations

import csv
import json
import os
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HOST = "127.0.0.1"
PORT = 8765
API_PATH = "/api/capstone-checkout"

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUTPUT_DIR = os.path.join(BASE_DIR, "Output files")
EVENT_LOG_CSV = os.path.join(OUTPUT_DIR, "checkout_events.csv")
STATE_CSV = os.path.join(OUTPUT_DIR, "checkout_current.csv")

EVENT_FIELDS = [
    "receivedAt",
    "articleKey",
    "articleUrl",
    "articleTitle",
    "checkedOutBy",
    "updatedAt",
    "clientIp",
]

STATE_FIELDS = [
    "articleKey",
    "articleUrl",
    "articleTitle",
    "checkedOutBy",
    "updatedAt",
    "receivedAt",
]


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def ensure_csv_with_header(path: str, fieldnames: list[str]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()


def load_state() -> dict[str, dict[str, str]]:
    state: dict[str, dict[str, str]] = {}
    if not os.path.exists(STATE_CSV):
        return state

    with open(STATE_CSV, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = (row.get("articleKey") or "").strip()
            if key:
                state[key] = row
    return state


def save_state(state: dict[str, dict[str, str]]) -> None:
    with open(STATE_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=STATE_FIELDS)
        writer.writeheader()
        for key in sorted(state):
            writer.writerow(state[key])


def normalize_payload(data: dict) -> dict[str, str]:
    return {
        "articleKey": str(data.get("articleKey", "")).strip(),
        "articleUrl": str(data.get("articleUrl", "")).strip(),
        "articleTitle": str(data.get("articleTitle", "")).strip(),
        "checkedOutBy": str(data.get("checkedOutBy", "")).strip(),
        "updatedAt": str(data.get("updatedAt", "")).strip(),
    }


def append_event(payload: dict[str, str], received_at: str, client_ip: str) -> None:
    row = {
        "receivedAt": received_at,
        "articleKey": payload["articleKey"],
        "articleUrl": payload["articleUrl"],
        "articleTitle": payload["articleTitle"],
        "checkedOutBy": payload["checkedOutBy"],
        "updatedAt": payload["updatedAt"],
        "clientIp": client_ip,
    }
    with open(EVENT_LOG_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=EVENT_FIELDS)
        writer.writerow(row)


def update_current_state(payload: dict[str, str], received_at: str) -> None:
    state = load_state()
    key = payload["articleKey"]

    # Empty checkedOutBy means the item was checked back in; remove from current state.
    if not payload["checkedOutBy"]:
        if key in state:
            del state[key]
        save_state(state)
        return

    state[key] = {
        "articleKey": key,
        "articleUrl": payload["articleUrl"],
        "articleTitle": payload["articleTitle"],
        "checkedOutBy": payload["checkedOutBy"],
        "updatedAt": payload["updatedAt"],
        "receivedAt": received_at,
    }
    save_state(state)


class CheckoutHandler(BaseHTTPRequestHandler):
    server_version = "CapstoneCheckoutHTTP/1.0"

    def _set_cors_headers(self) -> None:
        # Allow file:// pages (Origin: null) and local browser clients.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send_json(self, status_code: int, payload: dict) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status_code)
        self._set_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        if urlparse(self.path).path != API_PATH:
            self._send_json(404, {"ok": False, "error": "Not found"})
            return
        self.send_response(204)
        self._set_cors_headers()
        self.end_headers()

    def do_POST(self) -> None:
        if urlparse(self.path).path != API_PATH:
            self._send_json(404, {"ok": False, "error": "Not found"})
            return

        content_type = (self.headers.get("Content-Type") or "").lower()
        if "application/json" not in content_type:
            self._send_json(415, {"ok": False, "error": "Content-Type must be application/json"})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self._send_json(400, {"ok": False, "error": "Invalid Content-Length"})
            return

        try:
            raw = self.rfile.read(content_length)
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._send_json(400, {"ok": False, "error": "Invalid JSON body"})
            return

        if not isinstance(data, dict):
            self._send_json(400, {"ok": False, "error": "JSON body must be an object"})
            return

        payload = normalize_payload(data)
        if not payload["articleKey"]:
            self._send_json(400, {"ok": False, "error": "articleKey is required"})
            return

        received_at = utc_now_iso()
        client_ip = self.client_address[0] if self.client_address else ""

        append_event(payload, received_at, client_ip)
        update_current_state(payload, received_at)

        self._send_json(
            200,
            {
                "ok": True,
                "message": "Checkout update saved",
                "articleKey": payload["articleKey"],
            },
        )

    def log_message(self, format: str, *args) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {self.address_string()} - {format % args}")


def main() -> None:
    ensure_csv_with_header(EVENT_LOG_CSV, EVENT_FIELDS)
    ensure_csv_with_header(STATE_CSV, STATE_FIELDS)

    server = ThreadingHTTPServer((HOST, PORT), CheckoutHandler)
    print(f"Capstone checkout API running at http://{HOST}:{PORT}{API_PATH}")
    print(f"Event log CSV: {EVENT_LOG_CSV}")
    print(f"Current state CSV: {STATE_CSV}")
    print("Press Ctrl+C to stop.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
