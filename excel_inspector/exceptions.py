"""Domain exceptions for the Excel Structure Inspector (spec §4.1, §9).

The loader translates low-level openpyxl/zip failures (``BadZipFile``,
``InvalidFileException``, ...) into these explicit domain exceptions so the
rest of the pipeline can reason about them and stop early. Low-confidence or
undecidable analysis states are NOT exceptions; they are expressed as values
(``None`` + confidence, or accumulated ``warnings``) — see spec §6.
"""

from __future__ import annotations


class InspectorError(Exception):
    """Base class for all inspector domain errors."""


class CorruptWorkbookError(InspectorError):
    """Raised when a workbook cannot be opened due to corruption.

    Maps low-level failures such as ``zipfile.BadZipFile`` or openpyxl's
    ``InvalidFileException`` for damaged/non-.xlsx files (spec §9).
    """


class EncryptedWorkbookError(InspectorError):
    """Raised when a workbook is an OLE2/CFB container rather than OOXML zip.

    openpyxl alone only partially detects encryption, so the loader surfaces
    this explicitly to guide the user (spec §4.1, §9). The OLE2 ``d0cf11e0``
    magic is shared by password/encryption-protected ``.xlsx`` files *and* by
    unsupported legacy binary ``.xls`` files, so the classification carries
    that ambiguity; the message states it explicitly.
    """
