"""Normalize mobile contacts consistently across campaign and activity files."""

from __future__ import annotations

import math
import re
from decimal import Decimal, InvalidOperation
from typing import Any


PHONE_NORMALIZATION_VERSION = 1


def validated_phone(value: Any) -> str | None:
    """Keep the rightmost ten digits only when there are ten and they start at 9."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (not math.isfinite(value) or not value.is_integer()):
            return None
        text = str(int(value))
    elif isinstance(value, str):
        text = value.strip()
        if re.fullmatch(r"\+?\d+(?:\.\d+)?[eE][+-]?\d{1,3}|\+?\d+\.0+", text):
            try:
                number = Decimal(text)
                if number != number.to_integral_value() or not 0 <= number < 10**30:
                    return None
                text = str(int(number))
            except InvalidOperation:
                return None
    else:
        return None
    if not re.fullmatch(r"\+?[0-9(][0-9\s().-]*", text):
        return None
    digits = re.sub(r"[^0-9]", "", text)
    if len(digits) < 10:
        return None
    phone = digits[-10:]
    return phone if phone.startswith("9") else None


def first_valid_phone(value: Any) -> str | None:
    """Keep the first usable number when a source cell contains several contacts."""
    candidates = re.split(r"[;,|\r\n/]+", value) if isinstance(value, str) else (value,)
    return next((phone for candidate in candidates if (phone := validated_phone(candidate)) is not None), None)
