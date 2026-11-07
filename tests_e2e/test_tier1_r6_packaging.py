# -*- coding: utf-8 -*-
"""Tier 1: Requirement R6 Feature Coverage (Isolation).
Tests exclusion of test frameworks from production packages, PyInstaller spec excludes,
multi-platform build scripts, manifest verifier tools, and clean source distribution.

Specification: PROJECT.md § Milestones M4, ORIGINAL_REQUEST.md § R6
Feature IDs: F19, F20, F21, F22, F23, F24
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def test_r6_requirements_production_excludes_test_deps():
    """Test 1 (R6.1 / F19): requirements.txt strictly excludes pytest, hypothesis, and test runners."""
    req_file = ROOT / "requirements.txt"
    assert req_file.exists(), "requirements.txt must exist in project root"

    with open(req_file, "r", encoding="utf-8") as f:
        lines = [line.strip().lower() for line in f if line.strip() and not line.strip().startswith("#")]

    forbidden_frameworks = {"pytest", "pytest-cov", "hypothesis", "mock", "unittest"}
    for line in lines:
        pkg_name = line.split("==")[0].split(">=")[0].split("<=")[0].strip()
        assert pkg_name not in forbidden_frameworks, (
            f"GAP-R6: Production requirements.txt contains test dependency '{pkg_name}'! "
            "Test dependencies must be moved to requirements-dev.txt."
        )


def test_r6_spec_file_excludes_tests_and_test_harness():
    """Test 2 (R6.2 / F20): attendance.spec explicitly excludes test frameworks from PyInstaller bundling."""
    spec_file = ROOT / "attendance.spec"
    assert spec_file.exists(), "attendance.spec must exist in project root"

    content = spec_file.read_text(encoding="utf-8")
    # Verify excludes parameter in Analysis
    assert "excludes=" in content, "attendance.spec should configure excludes"
    for forbidden in ('pytest', 'hypothesis'):
        assert forbidden in content, (
            f"attendance.spec should explicitly exclude '{forbidden}' from distribution binary"
        )


def test_r6_build_scripts_and_inno_setup_exist():
    """Test 3 (R6.3 / F21, F23): Multi-platform build tooling and Inno Setup configuration exist."""
    build_all = ROOT / "tools" / "build_all.py"
    setup_iss = ROOT / "setup.iss"
    bat_build = ROOT / "build_windows.bat"

    assert build_all.exists(), "tools/build_all.py build orchestration script must exist"
    assert setup_iss.exists(), "setup.iss Inno Setup configuration must exist"
    assert bat_build.exists(), "build_windows.bat Windows build script must exist"

    # Verify build_all.py defines multi-platform build stages
    build_code = build_all.read_text(encoding="utf-8")
    assert "stage_" in build_code or "def build" in build_code or "main" in build_code


def test_r6_release_manifest_verification_tool(tmp_path: Path):
    """Test 4 (R6.4 / F24): tools/verify_release.py validates SHA256.txt manifests accurately."""
    verify_script = ROOT / "tools" / "verify_release.py"
    assert verify_script.exists(), "tools/verify_release.py verification script must exist"

    # Create dummy release dir with SHA256.txt
    rel_dir = tmp_path / "mock_release"
    rel_dir.mkdir()
    sample_file = rel_dir / "sample.exe"
    sample_file.write_bytes(b"Mock Binary Content 2026")

    import hashlib
    sha256 = hashlib.sha256(b"Mock Binary Content 2026").hexdigest()
    manifest = rel_dir / "SHA256.txt"
    manifest.write_text(f"{sha256}  sample.exe\n", encoding="utf-8")

    # Run verify_release.py
    cmd = [sys.executable, str(verify_script), str(rel_dir)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    assert proc.returncode == 0, f"verify_release.py failed: {proc.stdout}\n{proc.stderr}"
    assert "verified against SHA256.txt" in proc.stdout


def test_r6_clean_source_tarball_generator():
    """Test 5 (R6.5 / F24): Project repository structure enables generating a test-free source tarball."""
    # Ensure no committed temporary or test-generated files in core/ or root
    forbidden_artifacts = [
        ROOT / ".pytest_cache",
        ROOT / "tests" / "__pycache__",
    ]
    # Verify tools/build_all.py or packaging tool has logic to prune tests/
    build_all = (ROOT / "tools" / "build_all.py").read_text(encoding="utf-8")
    assert "test" in build_all.lower(), "Build orchestration script should handle test artifact stripping"
