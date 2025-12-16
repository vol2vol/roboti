import os
import ssl
import smtplib
import imaplib
import email
from email.message import EmailMessage
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime, formatdate
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
import mimetypes
from pathlib import Path

SMTP_SERVER = "smtp.mail.ru"
SMTP_PORT_SSL = 465

IMAP_SERVER = "imap.mail.ru"
IMAP_PORT_SSL = 993

OUTGOING_TO = ["vovalazar04@yandex.ru"]      # получатель(и) исходного письма
OUTGOING_SUBJECT = "Отчёт за неделю (автоматически)"
OUTGOING_BODY = (
    "Добрый день!\n\n"
    "Высылаю отчёт за неделю во вложении.\n"
    "Письмо отправлено автоматически скриптом.\n\n"
    "С уважением,\nАвтобот"
)
ATTACHMENT_PATH = r"C:\Users\vol\Desktop\учёба\роботы\sample.pdf"

# ПАРАМЕТРЫ ПОИСКА ВХОДЯЩИХ (непрочитанные за 24 часа)
MAILBOX = "INBOX"
LOOKBACK_HOURS = 24

# Адрес, на который нужно переслать найденные письма
FORWARD_TO = "vovalazar2004@yandex.ru"


def load_credentials():
    load_dotenv()
    email_addr = os.getenv("MAILRU_EMAIL")
    password = os.getenv("MAILRU_PASSWORD")
    if not email_addr or not password:
        raise RuntimeError("В .env должны быть заданы MAILRU_EMAIL и MAILRU_PASSWORD")
    return email_addr, password


def attach_file(msg: EmailMessage, filepath: str):
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Вложение не найдено: {path.resolve()}")
    ctype, encoding = mimetypes.guess_type(str(path))
    if ctype is None or encoding is not None:
        # Бинарный файл по умолчанию
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)

    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name
        )


def send_email_with_attachment(smtp_server, smtp_port, email_addr, password,
                               to_addrs, subject, body, attachment_path):
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=ctx) as server:
        server.login(email_addr, password)
        msg = EmailMessage()
        msg["From"] = email_addr
        msg["To"] = ", ".join(to_addrs)
        msg["Date"] = formatdate(localtime=True)
        msg["Subject"] = subject
        msg.set_content(body)

        attach_file(msg, attachment_path)

        server.send_message(msg)
        print(f"[OK] Отправлено письмо с вложением -> {to_addrs}")


def _decode_header_value(value):
    if value is None:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def _is_within_last_hours(dt: datetime, hours: int) -> bool:
    if dt.tzinfo is None:
        # INTERNALDATE обычно в локальном времени сервера; приводим к UTC для честного сравнения
        dt = dt.replace(tzinfo=timezone.utc)
    threshold = datetime.now(timezone.utc) - timedelta(hours=hours)
    return dt >= threshold


def forward_raw_message_via_smtp(raw_bytes: bytes, original_subject: str,
                                 smtp_server, smtp_port, email_addr, password, forward_to):
    """Пересылаем исходное письмо как вложение message/rfc822 (самый надёжный способ)."""
    fwd = EmailMessage()
    fwd["From"] = email_addr
    fwd["To"] = forward_to
    fwd["Subject"] = f"Fwd: {original_subject or '(без темы)'}"
    fwd["Date"] = formatdate(localtime=True)
    fwd.set_content(
        "Автоматическая пересылка непрочитанного письма, полученного за последние 24 часа. "
        "Исходное письмо приложено .eml вложением (message/rfc822)."
    )
    fwd.add_attachment(raw_bytes, maintype="message", subtype="rfc822", filename="forwarded.eml")

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=ctx) as server:
        server.login(email_addr, password)
        server.send_message(fwd)


def search_unseen_last_24h_and_forward(imap_server, imap_port, email_addr, password,
                                       mailbox, hours_to_look_back, forward_to):
    with imaplib.IMAP4_SSL(imap_server, imap_port) as imap:
        imap.login(email_addr, password)
        imap.select(mailbox)

        # 1) Сначала находим все непрочитанные
        status, data = imap.search(None, 'UNSEEN')
        if status != "OK":
            raise RuntimeError("Не удалось выполнить IMAP SEARCH UNSEEN")

        ids = data[0].split()
        if not ids:
            print("[INFO] Непрочитанных писем нет.")
            return

        forwarded = 0
        for msg_id in ids:
            # Получаем метаданные (INTERNALDATE) и само письмо
            status, fetch_data = imap.fetch(msg_id, "(RFC822 INTERNALDATE)")
            if status != "OK" or not fetch_data or fetch_data[0] is None:
                continue

            raw_bytes = None
            internaldate = None
            for part in fetch_data:
                if isinstance(part, tuple) and len(part) == 2:
                    # Часть с RFC822
                    if b"RFC822" in part[0]:
                        raw_bytes = part[1]
                if isinstance(part, bytes) and b"INTERNALDATE" in part:
                    # В некоторых реализациях приходит отдельной частью
                    pass

            if internaldate is None:
                status, id_data = imap.fetch(msg_id, "(INTERNALDATE)")
                if status == "OK" and id_data and id_data[0]:
                    meta = id_data[0]
                    if isinstance(meta, tuple):
                        meta = meta[0]
                    try:
                        internaldate_str = meta.decode(errors="ignore")
                        # Вынимаем строку даты в кавычках
                        start = internaldate_str.find('"')
                        end = internaldate_str.rfind('"')
                        if start != -1 and end != -1 and end > start:
                            internaldate_val = internaldate_str[start + 1:end]
                            internaldate = parsedate_to_datetime(internaldate_val)
                    except Exception:
                        internaldate = None

            if raw_bytes is None:
                # На некоторых серверах RFC822 приходит в первой части
                if fetch_data and isinstance(fetch_data[0], tuple):
                    raw_bytes = fetch_data[0][1]

            if internaldate is None:
                continue

            if not _is_within_last_hours(internaldate, hours_to_look_back):
                continue

            try:
                msg_obj = email.message_from_bytes(raw_bytes)
                orig_subject = _decode_header_value(msg_obj.get("Subject"))
            except Exception:
                orig_subject = ""

            try:
                forward_raw_message_via_smtp(
                    raw_bytes=raw_bytes,
                    original_subject=orig_subject,
                    smtp_server=SMTP_SERVER,
                    smtp_port=SMTP_PORT_SSL,
                    email_addr=email_addr,
                    password=password,
                    forward_to=forward_to
                )
                forwarded += 1
                print(f"[OK] Переслано письмо UID={msg_id.decode()} -> {forward_to}")
            except Exception as e:
                print(f"[ERR] Не удалось переслать UID={msg_id.decode()}: {e}")

        if forwarded == 0:
            print("[INFO] Непрочитанные письма найденные, но ни одно не попало в окно последних 24 часов.")
        else:
            print(f"[OK] Всего переслано: {forwarded}")


def main():
    email_addr, password = load_credentials()

    # 1) Отправляем письмо с вложением (PDF/Word)
    send_email_with_attachment(
        smtp_server=SMTP_SERVER,
        smtp_port=SMTP_PORT_SSL,
        email_addr=email_addr,
        password=password,
        to_addrs=OUTGOING_TO,
        subject=OUTGOING_SUBJECT,
        body=OUTGOING_BODY,
        attachment_path=ATTACHMENT_PATH
    )

    # 2) Находим непрочитанные за 24 часа и пересылаем их на FORWARD_TO
    search_unseen_last_24h_and_forward(
        imap_server=IMAP_SERVER,
        imap_port=IMAP_PORT_SSL,
        email_addr=email_addr,
        password=password,
        mailbox=MAILBOX,
        hours_to_look_back=LOOKBACK_HOURS,
        forward_to=FORWARD_TO
    )


if __name__ == "__main__":
    main()
