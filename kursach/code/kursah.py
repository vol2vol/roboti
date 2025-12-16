import imaplib
import email
import email.header
import email.utils
from email.message import Message
import json
import logging
import os
import re
import sqlite3
import time
from typing import Tuple, Optional

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError, expect

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

EMAIL_RE = re.compile(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", re.I)
PHONE_RE = re.compile(r"(\+?\d[\d\-\s\(\)]{8,}\d)")

DB_PATH = "rpa_dedup.sqlite3"
STATE_PATH = "yougile_state.json"


def decode_mime(value: str) -> str:
    if not value:
        return ""
    parts = email.header.decode_header(value)
    out = []
    for chunk, enc in parts:
        if isinstance(chunk, bytes):
            out.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(chunk)
    return "".join(out).strip()


def extract_body(msg: Message) -> Tuple[str, str]:
    """Return (content_type, body). Prefer text/plain else text/html."""
    if msg.is_multipart():
        plain = None
        html_part = None
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                continue

            payload = part.get_payload(decode=True)
            if payload is None:
                continue

            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")

            if ctype == "text/plain" and plain is None:
                plain = text
            elif ctype == "text/html" and html_part is None:
                html_part = text

        if plain is not None:
            return "text/plain", plain
        if html_part is not None:
            return "text/html", html_part
        return "text/plain", ""
    else:
        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        return msg.get_content_type(), payload.decode(charset, errors="replace")


def extract_contacts(from_header: str, body_text: str) -> str:
    name, addr = email.utils.parseaddr(from_header or "")
    emails = sorted(set(EMAIL_RE.findall(body_text or "")))
    phones = sorted(set(m.group(1) for m in PHONE_RE.finditer(body_text or "")))

    if addr and addr not in emails:
        emails.insert(0, addr)

    lines = []
    if name or addr:
        lines.append(f"Отправитель: {name} <{addr}>".strip())
    if emails:
        lines.append("Email: " + ", ".join(emails))
    if phones:
        lines.append("Телефон: " + ", ".join(phones))

    return "\n".join(lines).strip()


def build_description_plain(from_header: str, body: str) -> str:
    contacts = extract_contacts(from_header, body)
    return f"Контакты:\n{contacts}\n\n---\n\n{(body or '').strip()}".strip()


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


# ------------------ dedup store ------------------

class Dedup:
    def __init__(self, path: str):
        self.conn = sqlite3.connect(path)
        self.conn.execute(
            "CREATE TABLE IF NOT EXISTS processed ("
            " mailbox TEXT NOT NULL, "
            " message_id TEXT NOT NULL, "
            " ts INT NOT NULL, "
            " PRIMARY KEY(mailbox, message_id)"
            ")"
        )
        self.conn.commit()

    def seen(self, mailbox: str, message_id: str) -> bool:
        cur = self.conn.execute(
            "SELECT 1 FROM processed WHERE mailbox=? AND message_id=?",
            (mailbox, message_id),
        )
        return cur.fetchone() is not None

    def mark(self, mailbox: str, message_id: str):
        self.conn.execute(
            "INSERT OR REPLACE INTO processed(mailbox, message_id, ts) VALUES(?,?,?)",
            (mailbox, message_id, int(time.time())),
        )
        self.conn.commit()



def _any_frame_has(page, locator_str: str) -> bool:
    """Проверяет locator_str на page и во всех iframe."""
    try:
        if page.locator(locator_str).count() > 0:
            return True
    except Exception:
        pass

    for fr in page.frames:
        try:
            if fr.locator(locator_str).count() > 0:
                return True
        except Exception:
            continue
    return False


def wait_yougile_ready(page, timeout_ms: int = 60000) -> str:
    """
    Ждём, пока YouGile загрузит либо login-форму, либо интерфейс доски.
    Возвращает: "login" или "board".
    """
    start = time.time()

    login_markers = [
        "text=/Sign in/i",
        "input[placeholder*='e-mail' i]",
        "input[type='password']",
    ]

    board_markers = [
        "text=/\\+?\\s*Add task|\\+?\\s*Добавить/i",
        "text=/\\bBoard\\b|\\bДоска\\b/i",
        "text=/My Tasks|Мои задачи/i",
    ]

    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except PWTimeoutError:
        pass

    while (time.time() - start) * 1000 < timeout_ms:
        if any(_any_frame_has(page, m) for m in login_markers):
            return "login"
        if any(_any_frame_has(page, m) for m in board_markers):
            return "board"
        page.wait_for_timeout(250)

    page.screenshot(path="not_ready_timeout.png", full_page=True)
    raise RuntimeError("YouGile не успел загрузиться (ни login, ни board) за timeout. Скрин: not_ready_timeout.png")


def wait_board_ready(page, timeout_ms: int = 60000):
    """Ждём именно доску (а не логин)."""
    start = time.time()
    markers = [
        "text=/\\+?\\s*Add task|\\+?\\s*Добавить/i",
        "text=/\\bBoard\\b|\\bДоска\\b/i",
        "text=/My Tasks|Мои задачи/i",
    ]

    while (time.time() - start) * 1000 < timeout_ms:
        if any(_any_frame_has(page, m) for m in markers):
            return
        page.wait_for_timeout(250)

    page.screenshot(path="board_not_ready.png", full_page=True)
    raise RuntimeError("Доска не успела загрузиться за timeout. Скрин: board_not_ready.png")


def open_project(page, url: str) -> str:
    """
    Открываем страницу проекта/доски.
    Возвращает состояние: 'login' или 'board'
    """
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    return wait_yougile_ready(page, timeout_ms=60000)


def ensure_board_tab(page):
    # Если есть вкладка Board — кликнем (без ошибки, если её нет)
    try:
        tab = page.get_by_role("tab", name=re.compile(r"board|доска", re.I))
        if tab.count() > 0:
            tab.first.click(timeout=5000)
            page.wait_for_timeout(300)
    except Exception:
        pass


def find_login_scope(page):
    """Возвращает page или frame, где реально находятся поля логина."""
    try:
        if page.locator("input[placeholder*='e-mail' i]").count() > 0:
            return page
    except Exception:
        pass

    for fr in page.frames:
        try:
            if fr.locator("input[placeholder*='e-mail' i]").count() > 0:
                return fr
        except Exception:
            continue

    return page


def click_login(page, scope):
    candidates = [
        scope.locator("form button:has-text('Sign in')"),
        scope.locator("button:has-text('Sign in')"),
        scope.get_by_role("button", name=re.compile(r"sign\s*in", re.I)),
        scope.locator("[role='button']:has-text('Sign in')"),
        scope.locator("input[type='submit'][value*='Sign' i]"),
        scope.locator("input[type='button'][value*='Sign' i]"),
    ]
    for c in candidates:
        try:
            if c.count() > 0:
                c.first.scroll_into_view_if_needed()
                c.first.click(timeout=7000, force=True)
                return
        except Exception:
            continue

    # fallback
    try:
        page.keyboard.press("Enter")
    except Exception:
        pass


def yougile_login(page, email_: str, password_: str):
    page.wait_for_load_state("domcontentloaded")
    scope = find_login_scope(page)

    email_input = scope.locator("input[placeholder*='e-mail' i]")
    pass_input = scope.locator("input[type='password'], input[placeholder*='Password' i]")

    email_input.first.wait_for(state="visible", timeout=15000)
    pass_input.first.wait_for(state="visible", timeout=15000)

    expect(email_input.first).to_be_editable(timeout=15000)
    expect(pass_input.first).to_be_editable(timeout=15000)

    email_input.first.click()
    email_input.first.fill("")
    email_input.first.type(email_, delay=25)

    pass_input.first.click()
    pass_input.first.fill("")
    pass_input.first.type(password_, delay=25)

    click_login(page, scope)

    # Успех: появилось "My Tasks" / "Мои задачи"
    ok = page.locator("text=/My Tasks|Мои задачи/i")
    try:
        ok.wait_for(timeout=20000)
    except PWTimeoutError:
        page.screenshot(path="login_failed.png", full_page=True)
        raise RuntimeError(f"Не дождался успешного входа. URL={page.url}. Скрин: login_failed.png")


def find_board_scope(page):
    """
    Возвращает scope (page или frame), где реально видны элементы доски ('Add task').
    """
    mark = "text=/\\+?\\s*Add task|\\+?\\s*Добавить/i"
    try:
        if page.locator(mark).count() > 0:
            return page
    except Exception:
        pass

    for fr in page.frames:
        try:
            if fr.locator(mark).count() > 0:
                return fr
        except Exception:
            continue

    return page


def find_column_container(scope, column_name: str):
    """
    Пытаемся найти колонку по заголовку.
    Если не нашли — возвращаем первую колонку, где есть '+ Add task'.
    """
    target = normalize_spaces(column_name)
    pattern = re.compile(r"\s+".join(map(re.escape, target.split())), re.I)

    # 1) По имени колонки
    try:
        header = scope.get_by_text(pattern).first
        if header.count() > 0:
            header.scroll_into_view_if_needed()
            for lvl in range(1, 18):
                anc = header.locator(f"xpath=ancestor::*[self::div or self::section][{lvl}]")
                try:
                    if anc.count() == 0:
                        continue
                    if anc.locator("text=/\\+?\\s*Add task|\\+?\\s*Добавить/i").count() > 0:
                        return anc
                except Exception:
                    continue
            return header.locator("xpath=ancestor::*[self::div or self::section][1]")
    except Exception:
        pass

    # 2) Fallback: первая колонка с '+ Add task'
    add = scope.locator("text=/\\+?\\s*Add task|\\+?\\s*Добавить/i").first
    if add.count() > 0:
        add.scroll_into_view_if_needed()
        for lvl in range(1, 18):
            anc = add.locator(f"xpath=ancestor::*[self::div or self::section][{lvl}]")
            try:
                if anc.count() == 0:
                    continue
                if anc.locator("button, [role='button'], a").count() > 0:
                    logging.warning("Колонка '%s' не найдена. Использую первую колонку с '+ Add task'.", column_name)
                    return anc
            except Exception:
                continue
        return add.locator("xpath=ancestor::*[self::div or self::section][1]")

    return None


def create_task_ui(page, column_name: str, title: str, description: str):
    # ждём полной загрузки доски (исправляет вашу проблему со спиннером)
    wait_board_ready(page, timeout_ms=60000)

    ensure_board_tab(page)

    scope = find_board_scope(page)
    col = find_column_container(scope, column_name)

    if col is None or col.count() == 0:
        page.screenshot(path="ui_failed.png", full_page=True)
        raise RuntimeError(
            f"Не нашёл колонку '{column_name}' и не нашёл ни одной '+ Add task'. Скрин: ui_failed.png"
        )

    # нажать "+ Add task"
    candidates = [
        col.locator("text=/\\+\\s*Add task|Add task/i"),
        col.locator("text=/\\+\\s*Добавить|Добавить задачу/i"),
        col.get_by_role("button", name=re.compile(r"add task|добавить", re.I)),
        col.locator("[role='button']:has-text('Add task')"),
    ]
    clicked = False
    for c in candidates:
        try:
            if c.count() > 0:
                c.first.scroll_into_view_if_needed()
                c.first.click(timeout=5000, force=True)
                clicked = True
                break
        except Exception:
            continue

    if not clicked:
        page.screenshot(path="ui_failed.png", full_page=True)
        raise RuntimeError("Не смог нажать '+ Add task'. Скрин: ui_failed.png")

    page.wait_for_timeout(250)

    # ввод названия задачи
    input_candidates = [
        col.locator("input").first,
        col.locator("textarea").first,
        page.locator("input:focus").first,
        page.locator("textarea:focus").first,
        page.locator("[contenteditable='true']:focus").first,
    ]

    typed = False
    for ic in input_candidates:
        try:
            if ic.count() == 0:
                continue
            ic.click(timeout=2000)
            try:
                ic.fill("")
                ic.type(title, delay=15)
            except Exception:
                page.keyboard.press("Control+A")
                page.keyboard.insert_text(title)
            page.keyboard.press("Enter")
            typed = True
            break
        except Exception:
            continue

    if not typed:
        page.keyboard.insert_text(title)
        page.keyboard.press("Enter")

    # открыть карточку по названию
    card = scope.get_by_text(title).first
    try:
        card.wait_for(timeout=10000)
        card.click(timeout=5000)
    except Exception:
        pass

    page.wait_for_timeout(400)

    # вставить описание
    desc_candidates = [
        page.locator("textarea").first,
        page.get_by_placeholder(re.compile(r"description|описание", re.I)).first,
        page.locator("[contenteditable='true']").nth(0),
        page.locator("[contenteditable='true']").nth(1),
    ]
    for dc in desc_candidates:
        try:
            if dc.count() == 0 or not dc.is_visible():
                continue
            dc.click(timeout=2000)
            try:
                dc.fill(description)
            except Exception:
                page.keyboard.press("Control+A")
                page.keyboard.insert_text(description)
            break
        except Exception:
            continue

    page.keyboard.press("Escape")



def process_unseen(cfg: dict, dedup: Dedup, page):
    imap = imaplib.IMAP4_SSL(cfg["imap_host"], int(cfg.get("imap_port", 993)))
    imap.login(cfg["imap_user"], cfg["imap_password"])
    folder = cfg.get("imap_folder", "INBOX")
    imap.select(folder)

    typ, data = imap.search(None, "UNSEEN")
    if typ != "OK":
        imap.logout()
        return

    for num in data[0].split():
        typ, msg_data = imap.fetch(num, "(RFC822)")
        if typ != "OK":
            continue

        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        message_id = (msg.get("Message-ID") or "").strip()
        if not message_id:
            message_id = f"no-id:{num.decode(errors='ignore')}:{(msg.get('Date') or '').strip()}"

        if dedup.seen(folder, message_id):
            imap.store(num, "+FLAGS", "\\Seen")
            continue

        subject = decode_mime(msg.get("Subject", "")) or "(без темы)"
        from_header = decode_mime(msg.get("From", ""))

        _, body = extract_body(msg)
        description_plain = build_description_plain(from_header, body)

        try:
            state = open_project(page, cfg["board_url"])

            if state == "login":
                yougile_login(page, cfg["yougile_email"], cfg["yougile_password"])
                state = open_project(page, cfg["board_url"])
                page.context.storage_state(path=STATE_PATH)

            if state != "board":
                page.screenshot(path="ui_failed.png", full_page=True)
                raise RuntimeError("Не удалось попасть на доску после логина. Скрин: ui_failed.png")

            create_task_ui(page, cfg["column_name"], subject, description_plain)

            dedup.mark(folder, message_id)
            imap.store(num, "+FLAGS", "\\Seen")
            logging.info("Создана задача по письму: %s", subject)

        except Exception as e:
            logging.exception("Ошибка создания задачи по письму %s: %s", subject, e)
            # письмо НЕ помечаем Seen — попробует снова

    imap.logout()


def main():
    with open("config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)

    dedup = Dedup(DB_PATH)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=STATE_PATH) if os.path.exists(STATE_PATH) else browser.new_context()
        page = context.new_page()

        state = open_project(page, cfg["board_url"])
        if state == "login":
            logging.info("Нужна авторизация, выполняю вход...")
            yougile_login(page, cfg["yougile_email"], cfg["yougile_password"])
            state = open_project(page, cfg["board_url"])
            context.storage_state(path=STATE_PATH)
            logging.info("Сессия сохранена в %s", STATE_PATH)

        if state != "board":
            raise RuntimeError(f"Не удалось открыть доску. URL={page.url}")

        poll = int(cfg.get("poll_seconds", 30))

        try:
            while True:
                if page.is_closed():
                    logging.warning("Окно браузера закрыто. Открываю новое...")
                    page = context.new_page()
                    open_project(page, cfg["board_url"])

                process_unseen(cfg, dedup, page)
                time.sleep(poll)
        except KeyboardInterrupt:
            logging.info("Остановлено пользователем (Ctrl+C). Завершаю работу.")


if __name__ == "__main__":
    main()
