#!/usr/bin/env python3
"""AWS WorkMail IMAP/SMTP client.

This script uses WorkMail mailbox access credentials (username/password)
with standard IMAP and SMTP protocols to:
- list folders
- read messages
- send messages
- reply to a message

Required environment variables:
  WORKMAIL_USERNAME  WorkMail mailbox username (often full email address)
  WORKMAIL_PASSWORD  WorkMail mailbox password / app password

Optional environment variables:
  WORKMAIL_IMAP_HOST IMAP host (default: imap.mail.us-east-1.awsapps.com)
  WORKMAIL_IMAP_PORT IMAP SSL port (default: 993)
  WORKMAIL_SMTP_HOST SMTP host (default: smtp.mail.us-east-1.awsapps.com)
  WORKMAIL_SMTP_PORT SMTP STARTTLS port (default: 587)
  WORKMAIL_FROM      Default from address (default: WORKMAIL_USERNAME)
"""

from __future__ import annotations

import argparse
import email
import imaplib
import os
import smtplib
from dataclasses import dataclass
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from typing import Iterable


@dataclass
class WorkMailConfig:
    username: str
    password: str
    imap_host: str = "imap.mail.us-east-1.awsapps.com"
    imap_port: int = 993
    smtp_host: str = "smtp.mail.us-east-1.awsapps.com"
    smtp_port: int = 587
    from_address: str | None = None

    @classmethod
    def from_env(cls) -> "WorkMailConfig":
        username = os.environ.get("WORKMAIL_USERNAME")
        password = os.environ.get("WORKMAIL_PASSWORD")
        if not username or not password:
            raise ValueError(
                "WORKMAIL_USERNAME and WORKMAIL_PASSWORD must be set in the environment."
            )

        return cls(
            username=username,
            password=password,
            imap_host=os.environ.get("WORKMAIL_IMAP_HOST", cls.imap_host),
            imap_port=int(os.environ.get("WORKMAIL_IMAP_PORT", cls.imap_port)),
            smtp_host=os.environ.get("WORKMAIL_SMTP_HOST", cls.smtp_host),
            smtp_port=int(os.environ.get("WORKMAIL_SMTP_PORT", cls.smtp_port)),
            from_address=os.environ.get("WORKMAIL_FROM", username),
        )


class WorkMailClient:
    def __init__(self, config: WorkMailConfig):
        self.config = config

    def list_folders(self) -> list[str]:
        with imaplib.IMAP4_SSL(self.config.imap_host, self.config.imap_port) as imap:
            self._imap_login(imap)
            status, raw = imap.list()
            self._assert_ok(status, "Failed to list folders")
            return [line.decode("utf-8", errors="replace") for line in (raw or [])]

    def list_messages(self, folder: str = "INBOX", limit: int = 10) -> list[dict]:
        with imaplib.IMAP4_SSL(self.config.imap_host, self.config.imap_port) as imap:
            self._imap_login(imap)
            status, _ = imap.select(folder, readonly=True)
            self._assert_ok(status, f"Failed to select folder {folder!r}")
            status, data = imap.search(None, "ALL")
            self._assert_ok(status, "Failed to search messages")
            ids = data[0].split() if data and data[0] else []
            selected_ids = ids[-limit:]

            messages: list[dict] = []
            for msg_id in reversed(selected_ids):
                status, msg_data = imap.fetch(msg_id, "(RFC822.HEADER)")
                self._assert_ok(status, f"Failed to fetch message {msg_id.decode()}")
                header_bytes = self._extract_message_bytes(msg_data)
                msg = email.message_from_bytes(header_bytes)
                messages.append(
                    {
                        "id": msg_id.decode(),
                        "from": self._decode_header(msg.get("From", "")),
                        "to": self._decode_header(msg.get("To", "")),
                        "subject": self._decode_header(msg.get("Subject", "")),
                        "date": self._decode_header(msg.get("Date", "")),
                        "message_id": self._decode_header(msg.get("Message-ID", "")),
                    }
                )

            return messages

    def read_message(self, message_id: str, folder: str = "INBOX") -> dict:
        with imaplib.IMAP4_SSL(self.config.imap_host, self.config.imap_port) as imap:
            self._imap_login(imap)
            status, _ = imap.select(folder, readonly=True)
            self._assert_ok(status, f"Failed to select folder {folder!r}")
            status, msg_data = imap.fetch(message_id, "(RFC822)")
            self._assert_ok(status, f"Failed to fetch message {message_id}")
            raw_message = self._extract_message_bytes(msg_data)
            message = email.message_from_bytes(raw_message)

            text_body = self._extract_text_body(message, subtype="plain")
            html_body = self._extract_text_body(message, subtype="html")

            return {
                "id": message_id,
                "from": self._decode_header(message.get("From", "")),
                "to": self._decode_header(message.get("To", "")),
                "subject": self._decode_header(message.get("Subject", "")),
                "date": self._decode_header(message.get("Date", "")),
                "message_id": self._decode_header(message.get("Message-ID", "")),
                "in_reply_to": self._decode_header(message.get("In-Reply-To", "")),
                "text_body": text_body,
                "html_body": html_body,
            }

    def send_email(
        self,
        to_addresses: Iterable[str],
        subject: str,
        body: str,
        cc_addresses: Iterable[str] | None = None,
        bcc_addresses: Iterable[str] | None = None,
        html_body: str | None = None,
        in_reply_to: str | None = None,
        references: str | None = None,
    ) -> str:
        to_list = [x.strip() for x in to_addresses if x.strip()]
        cc_list = [x.strip() for x in (cc_addresses or []) if x.strip()]
        bcc_list = [x.strip() for x in (bcc_addresses or []) if x.strip()]

        if not to_list:
            raise ValueError("At least one recipient is required in --to")

        message = EmailMessage()
        message["From"] = self.config.from_address or self.config.username
        message["To"] = ", ".join(to_list)
        if cc_list:
            message["Cc"] = ", ".join(cc_list)
        message["Subject"] = subject
        message["Date"] = formatdate(localtime=True)
        message["Message-ID"] = make_msgid()

        if in_reply_to:
            message["In-Reply-To"] = in_reply_to
        if references:
            message["References"] = references

        message.set_content(body)
        if html_body:
            message.add_alternative(html_body, subtype="html")

        recipients = to_list + cc_list + bcc_list

        with smtplib.SMTP(self.config.smtp_host, self.config.smtp_port) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(self.config.username, self.config.password)
            smtp.send_message(message, to_addrs=recipients)

        return str(message["Message-ID"])

    def reply_to_message(
        self,
        message_id: str,
        reply_body: str,
        folder: str = "INBOX",
        reply_all: bool = False,
    ) -> str:
        original = self.read_message(message_id=message_id, folder=folder)
        subject = original["subject"] or ""
        if not subject.lower().startswith("re:"):
            subject = f"Re: {subject}"

        to_recipients = [original["from"]]
        cc_recipients: list[str] = []

        if reply_all:
            parsed_to = [x.strip() for x in (original["to"] or "").split(",") if x.strip()]
            parsed_to = [x for x in parsed_to if self.config.from_address not in x]
            cc_recipients.extend(parsed_to)

        quoted = f"\n\nOn {original['date']}, {original['from']} wrote:\n"
        for line in (original.get("text_body") or "").splitlines():
            quoted += f"> {line}\n"

        composed_body = f"{reply_body}{quoted}"
        return self.send_email(
            to_addresses=to_recipients,
            cc_addresses=cc_recipients,
            subject=subject,
            body=composed_body,
            in_reply_to=original.get("message_id"),
            references=original.get("message_id"),
        )

    @staticmethod
    def _assert_ok(status: str, message: str) -> None:
        if status != "OK":
            raise RuntimeError(message)

    @staticmethod
    def _imap_login(imap: imaplib.IMAP4_SSL) -> None:
        status, _ = imap.login(os.environ["WORKMAIL_USERNAME"], os.environ["WORKMAIL_PASSWORD"])
        if status != "OK":
            raise RuntimeError("IMAP login failed")

    @staticmethod
    def _extract_message_bytes(msg_data) -> bytes:
        for part in msg_data:
            if isinstance(part, tuple) and len(part) > 1:
                return part[1]
        raise RuntimeError("No message content returned from IMAP server")

    @staticmethod
    def _decode_header(value: str) -> str:
        if not value:
            return ""
        return str(make_header(decode_header(value)))

    @classmethod
    def _extract_text_body(cls, msg: email.message.Message, subtype: str = "plain") -> str:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == f"text/{subtype}" and "attachment" not in str(
                    part.get("Content-Disposition", "")
                ).lower():
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    if payload is None:
                        continue
                    return payload.decode(charset, errors="replace")
            return ""

        payload = msg.get_payload(decode=True)
        if payload is None:
            return ""
        charset = msg.get_content_charset() or "utf-8"
        return payload.decode(charset, errors="replace")


def _split_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [x.strip() for x in value.split(",") if x.strip()]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="AWS WorkMail IMAP/SMTP email client")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("folders", help="List mailbox folders")

    list_cmd = subparsers.add_parser("list", help="List messages")
    list_cmd.add_argument("--folder", default="INBOX", help="Folder/mailbox name (default: INBOX)")
    list_cmd.add_argument("--limit", type=int, default=10, help="Max messages to list (default: 10)")

    read_cmd = subparsers.add_parser("read", help="Read one message by IMAP sequence id")
    read_cmd.add_argument("--id", required=True, help="IMAP message sequence id")
    read_cmd.add_argument("--folder", default="INBOX", help="Folder/mailbox name (default: INBOX)")

    send_cmd = subparsers.add_parser("send", help="Send a new email")
    send_cmd.add_argument("--to", required=True, help="Comma-separated recipients")
    send_cmd.add_argument("--cc", help="Comma-separated CC recipients")
    send_cmd.add_argument("--bcc", help="Comma-separated BCC recipients")
    send_cmd.add_argument("--subject", required=True, help="Email subject")
    send_cmd.add_argument("--body", required=True, help="Plain text body")
    send_cmd.add_argument("--html-body", help="Optional HTML body")

    reply_cmd = subparsers.add_parser("reply", help="Reply to an existing message")
    reply_cmd.add_argument("--id", required=True, help="IMAP message sequence id")
    reply_cmd.add_argument("--body", required=True, help="Reply body")
    reply_cmd.add_argument("--folder", default="INBOX", help="Folder/mailbox name (default: INBOX)")
    reply_cmd.add_argument("--reply-all", action="store_true", help="Reply-all behavior")

    return parser


def main() -> None:
    args = build_parser().parse_args()
    config = WorkMailConfig.from_env()
    client = WorkMailClient(config)

    if args.command == "folders":
        for folder in client.list_folders():
            print(folder)
        return

    if args.command == "list":
        messages = client.list_messages(folder=args.folder, limit=args.limit)
        for msg in messages:
            print(
                f"{msg['id']:>4} | {msg['date'][:31]:31} | {msg['from'][:35]:35} | {msg['subject']}"
            )
        return

    if args.command == "read":
        message = client.read_message(message_id=args.id, folder=args.folder)
        for key in ["id", "from", "to", "subject", "date", "message_id", "in_reply_to"]:
            print(f"{key}: {message.get(key, '')}")
        print("\n--- text body ---")
        print(message.get("text_body", ""))
        if message.get("html_body"):
            print("\n--- html body ---")
            print(message.get("html_body", ""))
        return

    if args.command == "send":
        msg_id = client.send_email(
            to_addresses=_split_csv(args.to),
            cc_addresses=_split_csv(args.cc),
            bcc_addresses=_split_csv(args.bcc),
            subject=args.subject,
            body=args.body,
            html_body=args.html_body,
        )
        print(f"Sent email with Message-ID: {msg_id}")
        return

    if args.command == "reply":
        msg_id = client.reply_to_message(
            message_id=args.id,
            reply_body=args.body,
            folder=args.folder,
            reply_all=args.reply_all,
        )
        print(f"Sent reply with Message-ID: {msg_id}")
        return

    raise RuntimeError(f"Unknown command: {args.command}")


if __name__ == "__main__":
    main()
