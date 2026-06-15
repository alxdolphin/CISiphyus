from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import parse_qsl, urlsplit

import pytest

import config
import report
import url_refresh


SAMPLE_LS = (
    "Procedure%3DspExportStudentMetrics%26"
    "%2540Enrollment_ProgramID%3D1330%26"
    "%2540CurrentUserID%3D12345"
)


def _sample_export_url(ls_blob: str = SAMPLE_LS) -> str:
    return f"https://cw.caseworthy.net/cis_prod.caseworthy/caseworthy/excelexport.aspx?FormID=99&LS={ls_blob}"


class TestSchoolYearConfig:
    def test_default_school_year_picks_latest(self):
        programs = {"SY23-24": 1200, "SY24-25": 1330, "SY25-26": 1331}
        assert config.default_school_year(programs) == "SY25-26"

    def test_default_school_year_env_override(self, monkeypatch):
        programs = {"SY24-25": 1330, "SY25-26": 1331}
        monkeypatch.setenv("CISDM_DEFAULT_SCHOOL_YEAR", "SY24-25")
        assert config.default_school_year(programs) == "SY24-25"

    def test_default_school_year_env_unknown_raises(self, monkeypatch):
        programs = {"SY25-26": 1331}
        monkeypatch.setenv("CISDM_DEFAULT_SCHOOL_YEAR", "SY99-00")
        with pytest.raises(ValueError, match="not in school_year_programs"):
            config.default_school_year(programs)

    def test_default_school_year_empty_map(self):
        assert config.default_school_year({}) is None

    def test_resolve_program_id_whitespace(self):
        programs = {"SY24-25": 1330}
        assert config.resolve_program_id(programs, "  SY24-25  ") == 1330

    def test_resolve_program_id_unknown(self):
        with pytest.raises(ValueError, match="Unknown school year"):
            config.resolve_program_id({"SY24-25": 1330}, "SY99-00")


class TestUrlRefresh:
    def test_replace_ls_param_round_trip(self):
        url = _sample_export_url()
        updated = url_refresh.replace_ls_param(
            url,
            "@Enrollment_ProgramID",
            "1331",
        )
        query = dict(parse_qsl(urlsplit(updated).query, keep_blank_values=True))
        procedure, pairs = url_refresh.decode_ls_pairs(query["LS"])
        assert procedure == "spExportStudentMetrics"
        enrollment = dict(pairs)["@Enrollment_ProgramID"]
        assert enrollment == "1331"

    def test_resolve_year_scoped_url(self):
        url = _sample_export_url()
        scoped = url_refresh.resolve_year_scoped_url(
            url,
            param="@Enrollment_ProgramID",
            program_id=1331,
        )
        assert scoped != url
        assert "1331" in scoped

    def test_replace_ls_param_missing_param_raises(self):
        url = _sample_export_url()
        with pytest.raises(ValueError, match="ls param not found"):
            url_refresh.replace_ls_param(url, "@MissingParam", "1")


class TestOutputPaths:
    def test_archives_when_non_default_year(self):
        output_dir, raw_path = report._resolve_output_paths(
            "student_metrics_summary",
            school_year="SY24-25",
            default_school_year="SY25-26",
        )
        assert output_dir == config.archives_pull_dir("SY24-25", "student_metrics_summary")
        assert raw_path == config.archives_raw_path("SY24-25", "student_metrics_summary")

    def test_latest_when_default_year(self):
        output_dir, raw_path = report._resolve_output_paths(
            "student_metrics_summary",
            school_year="SY25-26",
            default_school_year="SY25-26",
        )
        assert output_dir == config.latest_report_dir("student_metrics_summary")
        assert raw_path == config.latest_raw_path("student_metrics_summary")

    def test_latest_when_no_school_year(self):
        output_dir, raw_path = report._resolve_output_paths(
            "student_metrics_summary",
            school_year=None,
            default_school_year="SY25-26",
        )
        assert output_dir == config.latest_report_dir("student_metrics_summary")
        assert raw_path == config.latest_raw_path("student_metrics_summary")
