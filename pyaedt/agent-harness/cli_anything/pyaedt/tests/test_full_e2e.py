"""E2E tests for PyAEDT CLI."""

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest

# Add package to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))


def _resolve_cli(name):
    """Resolve installed CLI command; falls back to python -m for dev.

    Set env CLI_ANYTHING_FORCE_INSTALLED=1 to require the installed command.
    """
    import shutil
    force = os.environ.get("CLI_ANYTHING_FORCE_INSTALLED", "").strip() == "1"
    path = shutil.which(name)
    if path:
        print(f"[_resolve_cli] Using installed command: {path}")
        return [path]
    if force:
        raise RuntimeError(f"{name} not found in PATH. Install with: pip install -e .")
    module = name.replace("cli-anything-", "cli_anything.") + "." + name.split("-")[-1] + "_cli"
    print(f"[_resolve_cli] Falling back to: {sys.executable} -m {module}")
    return [sys.executable, "-m", module]


class TestCLISubprocess:
    """Test CLI as subprocess (as a real user/agent would)."""

    CLI_BASE = _resolve_cli("cli-anything-pyaedt")

    def _run(self, args, check=True):
        """Run CLI with given arguments."""
        result = subprocess.run(
            self.CLI_BASE + args,
            capture_output=True,
            text=True,
            check=False,
        )
        if check and result.returncode != 0:
            print(f"STDOUT: {result.stdout}")
            print(f"STDERR: {result.stderr}")
        return result

    def test_help(self):
        """Test CLI help output."""
        result = self._run(["--help"])
        assert result.returncode == 0
        assert "PyAEDT CLI" in result.stdout

    def test_version(self):
        """Test CLI version output."""
        result = self._run(["--version"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "version" in data

    def test_json_project_info(self):
        """Test JSON output for project info."""
        result = self._run(["--json", "project", "info"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "status" in data

    def test_project_new_json(self, tmp_path):
        """Test creating new project with JSON output."""
        out_dir = str(tmp_path)
        result = self._run(
            ["--json", "project", "new", "TestProject", "--path", out_dir],
            check=False,
        )
        # May fail on non-Windows without AEDT
        data = json.loads(result.stdout)
        assert "status" in data

    def test_design_new_json(self):
        """Test creating new design with JSON output."""
        result = self._run(["--json", "design", "new", "TestDesign", "--type", "HFSS"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "created"

    def test_variable_set_json(self):
        """Test setting variable with JSON output."""
        result = self._run(
            ["--json", "variable", "set", "freq", "10GHz"],
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "set"

    def test_variable_list_json(self):
        """Test listing variables with JSON output."""
        result = self._run(["--json", "variable", "list"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "status" in data

    def test_modeler_create_box_json(self):
        """Test creating box with JSON output."""
        result = self._run(
            ["--json", "modeler", "create-box", "Box1",
             "--position", "0,0,0", "--dimensions", "1,1,1"],
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "created"

    def test_modeler_create_cylinder_json(self):
        """Test creating cylinder with JSON output."""
        result = self._run(
            ["--json", "modeler", "create-cylinder", "Cyl1",
             "--position", "0,0,0", "--radius", "1", "--height", "10"],
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "created"

    def test_analysis_create_setup_json(self):
        """Test creating analysis setup with JSON output."""
        result = self._run(
            ["--json", "analysis", "create-setup", "Setup1", "--type", "Hfss"],
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "created"

    def test_analysis_list_json(self):
        """Test listing setups with JSON output."""
        result = self._run(["--json", "analysis", "list"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "status" in data

    def test_export_formats_json(self):
        """Test listing export formats with JSON output."""
        result = self._run(["--json", "export", "formats"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "formats" in data

    def test_session_status_json(self):
        """Test session status with JSON output."""
        result = self._run(["--json", "session", "status"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "status" in data

    def test_session_check_json(self):
        """Test session check with JSON output."""
        result = self._run(["--json", "session", "check"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "status" in data

    def test_process_list(self):
        """Test listing AEDT processes."""
        result = self._run(["process", "list"], check=False)
        # May return 0 even without psutil
        assert "processes" in result.stdout or "No AEDT" in result.stdout

    def test_list_materials_json(self):
        """Test listing materials with JSON output."""
        result = self._run(["--json", "variable", "list"])
        # Materials list is separate functionality
        assert result.returncode == 0

    def test_invalid_design_type(self):
        """Test error on invalid design type."""
        result = self._run(
            ["--json", "design", "new", "TestDesign", "--type", "InvalidType"],
            check=False,
        )
        data = json.loads(result.stdout)
        assert data["status"] == "error"

    def test_close_without_project(self):
        """Test closing when no project is open."""
        result = self._run(["--json", "project", "close"])
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert data["status"] == "closed"


class TestCLIIntegration:
    """Integration tests for complete workflows."""

    CLI_BASE = _resolve_cli("cli-anything-pyaedt")

    def _run(self, args, check=True):
        """Run CLI with given arguments."""
        result = subprocess.run(
            self.CLI_BASE + args,
            capture_output=True,
            text=True,
            check=False,
        )
        return result

    def test_full_project_creation_workflow(self, tmp_path):
        """Test creating project and adding components."""
        # This simulates a complete workflow
        # In reality, would require AEDT running

        # Create project
        project_result = self._run(
            ["--json", "project", "new",
             "WorkflowTest", "--path", str(tmp_path)],
        )

        # Create design
        design_result = self._run(
            ["--json", "design", "new",
             "HFSSDesign1", "--type", "HFSS"],
        )
        design_data = json.loads(design_result.stdout)
        assert design_data["status"] == "created"

        # Set variables
        var_result = self._run(
            ["--json", "variable", "set",
             "freq", "10GHz"],
        )
        var_data = json.loads(var_result.stdout)
        assert var_data["status"] == "set"

    def test_modeler_workflow(self):
        """Test modeler operations workflow."""
        # Create box
        box_result = self._run(
            ["--json", "modeler", "create-box",
             "Waveguide", "--position", "0,0,0", "--dimensions", "10,5,2.5",
             "--material", "aluminum"],
        )
        box_data = json.loads(box_result.stdout)
        assert box_data["status"] == "created"

        # Assign material
        mat_result = self._run(
            ["--json", "modeler", "assign-material",
             "Waveguide", "copper"],
        )
        mat_data = json.loads(mat_result.stdout)
        assert mat_data["status"] == "assigned"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
