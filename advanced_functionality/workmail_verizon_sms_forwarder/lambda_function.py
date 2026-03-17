"""AWS Lambda handler for forwarding Amazon WorkMail messages as Verizon SMS.

This function is designed for an Amazon WorkMail *inbound message flow rule*.
When WorkMail invokes the Lambda, the function fetches the raw email content,
extracts key fields, and sends a compact SMS-formatted email through Amazon SES
to Verizon's SMTP-to-SMS gateway.
"""

from __future__ import annotations

import email
import logging
import os
import re
from email import policy
from typing import Any

import boto3

LOGGER = logging.getLogger()
LOGGER.setLevel(logging.INFO)

WORKMAIL_MSG_FLOW = boto3.client("workmailmessageflow")
SES = boto3.client("ses")

# Accept local US 10-digit numbers with optional +1 prefix.
PHONE_PATTERN = re.compile(r"^\+?1?(\d{10})$")


def _normalized_phone(raw_phone: str | None) -> str:
    if not raw_phone:
        raise ValueError("No phone number was supplied.")

    digits_only = re.sub(r"\D", "", raw_phone)
    match = PHONE_PATTERN.match(digits_only)
    if not match:
        raise ValueError(
            "Expected TARGET_PHONE to contain a US 10-digit phone number "
            "(optionally prefixed by 1)."
        )

    return match.group(1)


def _extract_text_body(msg: email.message.EmailMessage) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    return payload.decode(charset, errors="replace").strip()
        return ""

    payload = msg.get_payload(decode=True)
    if not payload:
        return ""

    charset = msg.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="replace").strip()


def _build_sms_body(parsed: email.message.EmailMessage) -> str:
    source = parsed.get("From", "Unknown sender")
    subject = parsed.get("Subject", "(no subject)")
    text = _extract_text_body(parsed)

    compact_text = " ".join(text.split())
    if len(compact_text) > 120:
        compact_text = f"{compact_text[:117]}..."

    message = f"From: {source}\nSubj: {subject}\n{compact_text}".strip()

    # Verizon SMS gateway has practical limits near 160 chars.
    return message[:160]


def lambda_handler(event: dict[str, Any], context: Any) -> dict[str, str]:
    LOGGER.info("Incoming WorkMail event: %s", event)

    message_id = event.get("messageId")
    if not message_id:
        raise ValueError("WorkMail event is missing required field: messageId")

    target_phone = _normalized_phone(os.environ.get("TARGET_PHONE"))
    recipient = f"{target_phone}@vtext.com"
    source_email = os.environ["SOURCE_EMAIL"]

    raw_msg = WORKMAIL_MSG_FLOW.get_raw_message_content(messageId=message_id)
    parsed = email.message_from_bytes(
        raw_msg["messageContent"].read(),
        policy=policy.default,
    )

    sms_body = _build_sms_body(parsed)
    SES.send_email(
        Source=source_email,
        Destination={"ToAddresses": [recipient]},
        Message={
            "Subject": {"Data": "WorkMail alert", "Charset": "UTF-8"},
            "Body": {"Text": {"Data": sms_body, "Charset": "UTF-8"}},
        },
    )

    return {"status": "ok", "recipient": recipient}
