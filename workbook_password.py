"""Open password-protected source workbooks in memory for the current upload."""

from __future__ import annotations

from io import BytesIO
from zipfile import is_zipfile

from automation import AutomationError


ENCRYPTED_WORKBOOK_VERSION = 1
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def needs_workbook_password(data: bytes) -> bool:
    return data.startswith(OLE_SIGNATURE)


def unlock_workbook(data: bytes, password: str) -> bytes:
    """Return ordinary XLSX bytes without changing the source or saving a password."""
    if not needs_workbook_password(data):
        return data
    if not password:
        raise AutomationError("Enter the workbook password to prepare the files.")
    import msoffcrypto
    from msoffcrypto import exceptions

    try:
        office = msoffcrypto.OfficeFile(BytesIO(data))
        if not office.is_encrypted():
            raise AutomationError("Save this source as an .xlsx workbook in Excel before uploading it.")
        office.load_key(password=password, verify_password=True)
        decrypted = BytesIO()
        office.decrypt(decrypted, verify_integrity=True)
        result = decrypted.getvalue()
        if not is_zipfile(BytesIO(result)):
            raise AutomationError("Save this source as an .xlsx workbook in Excel before uploading it.")
        return result
    except exceptions.InvalidKeyError as exc:
        raise AutomationError("The workbook password is incorrect. Try again.") from exc
    except (exceptions.FileFormatError, exceptions.ParseError, exceptions.DecryptionError, OSError, ValueError) as exc:
        raise AutomationError("This workbook could not be unlocked. Check its password or save an unencrypted .xlsx copy in Excel.") from exc
