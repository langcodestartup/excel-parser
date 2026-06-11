"""CLI tests (plan v2 Phase 13 Step 1, L8): ``python -m excel_inspector``.

Covers the plan §7 Step 1 contract:

* ``--format json`` -> stdout parses with :func:`json.loads` and carries
  ``schema_version`` (plus golden record values, so a fast-but-wrong dump
  cannot pass);
* markdown (default) -> the ``| --- |`` separator row is present and
  ``--max-rows`` truncates with the "... more rows" trailer;
* corrupt / encrypted input -> exit code != 0, **nothing on stdout**, and an
  explicit ``error:`` line on stderr;
* path errors (review LOW): a nonexistent path -> ``file not found``, a
  directory -> ``not a file`` (exit 1, never the misleading 'corrupt' error);
  a negative ``--max-rows`` -> argparse usage error (exit 2);
* the real ``python -m excel_inspector`` entry point round-trips through a
  subprocess (both the success and the failure exit codes), so the
  ``__main__`` wiring itself is exercised, not just :func:`main`.

In-process tests drive :func:`excel_inspector.__main__.main` directly and
capture output via capsys (plan v2 §7).
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

from excel_inspector.__main__ import main
from excel_inspector.results import SCHEMA_VERSION

#: Project root — ``python -m excel_inspector`` resolves the (uninstalled)
#: package from the working directory, exactly like a developer invocation.
_PROJECT_ROOT = Path(__file__).resolve().parents[1]


# ---------------------------------------------------------------------------
# JSON mode
# ---------------------------------------------------------------------------


def test_cli_json_parses_and_carries_schema_version(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """--format json: stdout is valid JSON with the v1.0 schema marker."""

    exit_code = main(
        [str(fixture_path("offset_plus_subtotals")), "--format", "json"]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.err == ""
    parsed = json.loads(captured.out)
    assert parsed["schema_version"] == SCHEMA_VERSION
    # Golden guard (fixture single source of truth): 6 data rows, sum 590,
    # and no subtotal/total label leaked into the records.
    (sheet,) = parsed["sheets"]
    (table,) = sheet["tables"]
    assert table["row_count"] == 6
    assert sum(record["amount"] for record in table["records"]) == 590


def test_cli_json_emits_every_row_regardless_of_max_rows(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """JSON mode ignores --max-rows (plan §7 Step 1: a full extraction)."""

    exit_code = main(
        [str(fixture_path("header_simple")), "--format", "json", "--max-rows", "2"]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    table = json.loads(captured.out)["sheets"][0]["tables"][0]
    assert table["row_count"] == 5
    assert len(table["records"]) == 5  # all rows, not 2


# ---------------------------------------------------------------------------
# Markdown mode (the default)
# ---------------------------------------------------------------------------


def test_cli_markdown_default_has_separator_row(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """Default (markdown) output renders the table with a | --- | separator."""

    exit_code = main([str(fixture_path("header_simple"))])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.err == ""
    lines = captured.out.splitlines()
    assert lines[0] == "### Sheet1!T1"
    header = next(line for line in lines if line.startswith("|"))
    assert "name" in header
    separator = lines[lines.index(header) + 1]
    assert set(separator.replace("|", "").replace(" ", "")) == {"-"}


def test_cli_markdown_max_rows_truncates(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """--max-rows bounds the rendered rows with the '... more rows' trailer."""

    exit_code = main([str(fixture_path("header_simple")), "--max-rows", "2"])
    captured = capsys.readouterr()

    assert exit_code == 0
    # header_simple has 5 data rows; 2 rendered -> 3 announced as remaining.
    assert "3 more rows" in captured.out


def test_cli_output_is_deterministic(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """Two identical CLI invocations print byte-identical JSON (plan §8)."""

    path = str(fixture_path("multi_table_stacked"))
    assert main([path, "--format", "json"]) == 0
    first = capsys.readouterr().out
    assert main([path, "--format", "json"]) == 0
    second = capsys.readouterr().out
    assert first == second


# ---------------------------------------------------------------------------
# Error paths: corrupt / encrypted -> exit code != 0 + explicit stderr
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("fixture_id", "needle"),
    [
        ("corrupt", "corrupt"),
        ("encrypted", "password"),
    ],
)
def test_cli_unopenable_file_fails_with_stderr(
    fixture_path,
    capsys: pytest.CaptureFixture[str],
    fixture_id: str,
    needle: str,
) -> None:
    """Corrupt/encrypted input -> non-zero exit, empty stdout, stderr error."""

    exit_code = main([str(fixture_path(fixture_id)), "--format", "json"])
    captured = capsys.readouterr()

    assert exit_code != 0
    assert captured.out == ""  # a failure must never look like empty output
    assert captured.err.startswith("error: ")
    assert needle.lower() in captured.err.lower()


def test_cli_negative_max_rows_rejected_as_usage_error(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """--max-rows -1 -> argparse usage error: exit code 2, nothing on stdout.

    Review LOW: a negative budget would reach ``DataFrame.head(-n)``, which
    silently *drops* the last n rows instead of failing — so the value is
    rejected at parse time with an explicit message.
    """

    with pytest.raises(SystemExit) as excinfo:
        main([str(fixture_path("header_simple")), "--max-rows", "-1"])
    captured = capsys.readouterr()

    assert excinfo.value.code == 2
    assert captured.out == ""
    assert "--max-rows" in captured.err
    assert "non-negative" in captured.err


def test_cli_max_rows_zero_is_accepted(
    fixture_path, capsys: pytest.CaptureFixture[str]
) -> None:
    """--max-rows 0 stays valid (only negatives are rejected): all truncated."""

    exit_code = main([str(fixture_path("header_simple")), "--max-rows", "0"])
    captured = capsys.readouterr()

    assert exit_code == 0
    # header_simple has 5 data rows; 0 rendered -> all 5 announced remaining.
    assert "5 more rows" in captured.out


def test_cli_missing_path_reports_file_not_found(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    """A nonexistent path -> 'file not found', never a 'corrupt' error.

    Review LOW: both a missing path and a directory used to fall through to
    openpyxl and surface as 'corrupt or not a valid .xlsx'.
    """

    missing = tmp_path / "nope.xlsx"
    exit_code = main([str(missing), "--format", "json"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert captured.out == ""
    assert captured.err.startswith("error: ")
    assert "file not found" in captured.err
    assert str(missing) in captured.err
    assert "corrupt" not in captured.err


def test_cli_directory_path_reports_not_a_file(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    """A directory path -> 'not a file' (distinct from 'file not found')."""

    exit_code = main([str(tmp_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert captured.out == ""
    assert captured.err.startswith("error: ")
    assert "not a file" in captured.err
    assert str(tmp_path) in captured.err
    assert "file not found" not in captured.err
    assert "corrupt" not in captured.err


# ---------------------------------------------------------------------------
# The real module entry point (subprocess round-trip)
# ---------------------------------------------------------------------------


def test_module_entrypoint_subprocess_success(fixture_path) -> None:
    """``python -m excel_inspector <file> --format json`` exits 0 with JSON."""

    proc = subprocess.run(
        [
            sys.executable,
            "-m",
            "excel_inspector",
            str(fixture_path("header_simple")),
            "--format",
            "json",
        ],
        capture_output=True,
        text=True,
        check=False,
        cwd=_PROJECT_ROOT,
    )

    assert proc.returncode == 0, proc.stderr
    assert json.loads(proc.stdout)["schema_version"] == SCHEMA_VERSION


def test_module_entrypoint_subprocess_corrupt_exits_nonzero(
    fixture_path,
) -> None:
    """The corrupt-file failure propagates as a real process exit code."""

    proc = subprocess.run(
        [sys.executable, "-m", "excel_inspector", str(fixture_path("corrupt"))],
        capture_output=True,
        text=True,
        check=False,
        cwd=_PROJECT_ROOT,
    )

    assert proc.returncode != 0
    assert proc.stdout == ""
    assert "error:" in proc.stderr
