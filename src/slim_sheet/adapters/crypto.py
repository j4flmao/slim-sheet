"""Password-protected workbooks (optional extra `crypto`, msoffcrypto-tool).

A password-protected xlsx is an OLE2 container holding `EncryptionInfo` and
`EncryptedPackage` streams. We detect that without the extra (so the error can say what
to install), decrypt into a private temporary folder, and delete it when the caller is
done. The password is read from SLIM_SHEET_PASSWORD only — never from a tool argument —
so it does not end up in chat transcripts.
"""

from __future__ import annotations

import shutil
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError

_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_MARKER = "EncryptionInfo".encode("utf-16-le")  # OLE directory entries are UTF-16 names
INSTALL_HINT = (
    "Install the crypto extra: `uv sync --extra crypto` or "
    "`uvx --from 'slim-sheet[crypto]' slim-sheet mcp`."
)


def is_encrypted(path: Path) -> bool:
    with path.open("rb") as handle:
        if handle.read(8) != _OLE2:
            return False
        handle.seek(0)
        return _MARKER in handle.read()


@contextmanager
def decrypted(path: Path, password: str | None) -> Iterator[Path]:
    """Yield a decrypted temporary copy of `path`; it is removed on exit."""
    if not password:
        raise SlimError(
            ErrorCode.ENCRYPTED,
            f"'{path.name}' is password-protected",
            hint="Set SLIM_SHEET_PASSWORD in the server's environment (not in chat) and retry.",
        )
    try:
        import msoffcrypto
        from msoffcrypto.exceptions import InvalidKeyError
    except ImportError as exc:
        raise SlimError(
            ErrorCode.MISSING_EXTRA, f"'{path.name}' is encrypted", hint=INSTALL_HINT
        ) from exc
    folder = Path(tempfile.mkdtemp(prefix="slim-decrypt-"))
    target = folder / (path.stem + ".xlsx")
    try:
        with path.open("rb") as source, target.open("wb") as out:
            office = msoffcrypto.OfficeFile(source)
            office.load_key(password=password)
            office.decrypt(out)
    except InvalidKeyError as exc:
        shutil.rmtree(folder, ignore_errors=True)
        raise SlimError(ErrorCode.ENCRYPTED, f"Wrong password for '{path.name}'") from exc
    except Exception as exc:
        shutil.rmtree(folder, ignore_errors=True)
        raise SlimError(ErrorCode.READ_FAILED, f"Could not decrypt '{path.name}': {exc}") from exc
    try:
        yield target
    finally:
        shutil.rmtree(folder, ignore_errors=True)
