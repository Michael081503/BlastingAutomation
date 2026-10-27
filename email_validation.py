"""Validate email syntax locally without checking or contacting mail servers."""

from __future__ import annotations

from typing import Any

from email_validator import EmailNotValidError, validate_email


EMAIL_VALIDATION_VERSION = 1


def validated_email(value: Any) -> str | None:
    """Return a normalized address, or None when it cannot be used for email.

    Only surrounding whitespace is removed. Internal spaces and invalid dots
    are rejected rather than guessed at. Valid syntax does not prove that the
    mailbox exists or that a message will be delivered.
    """
    if not isinstance(value, str) or not value.strip():
        return None
    try:
        return validate_email(value.strip(), check_deliverability=False, strict=True).normalized
    except EmailNotValidError:
        return None
