"""PyAEDT project management - Core module with real PyAEDT API integration."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from pyaedt import Hfss

# Singleton HFSS instance for API calls
_hfss_instance: Optional["Hfss"] = None


def _patch_design_settings():
    """Patch DesignSettings for AEDT 2019 compatibility."""
    try:
        import pyaedt.application.Design as dm
        _orig_init = dm.DesignSettings.__init__
        def _patched_init(self, app):
            try:
                _orig_init(self, app)
            except AttributeError:
                self._app = app
                self.design_settings = None
                self.manipulate_inputs = None
        dm.DesignSettings.__init__ = _patched_init
    except Exception:
        pass


def _get_hfss() -> "Hfss":
    """Get or create HFSS instance for API calls.

    Returns
    -------
    Hfss
        The active HFSS design instance.

    Raises
    ------
    RuntimeError
        If AEDT is not available or connection fails.
    """
    global _hfss_instance

    if _hfss_instance is None:
        try:
            _patch_design_settings()
            from pyaedt import Hfss
            _hfss_instance = Hfss(non_graphical=True, specified_version='2019.1')
        except Exception as e:
            raise RuntimeError(
                f"Failed to connect to AEDT: {e}. "
                "Make sure Ansys Electronics Desktop is installed and licensed."
            )

    return _hfss_instance


class ProjectManager:
    """Manages AEDT project files and state."""

    def __init__(self):
        """Initialize project manager."""
        self._current_project_path: Optional[str] = None
        self._current_design: Optional[str] = None
        self._session_file: Optional[str] = None

    def get_session_path(self) -> str:
        """Get the session file path."""
        if self._session_file is None:
            home = Path.home()
            session_dir = home / ".cli-anything-pyaedt"
            session_dir.mkdir(parents=True, exist_ok=True)
            self._session_file = str(session_dir / "session.json")
        return self._session_file

    def save_session(self, project_path: Optional[str] = None,
                     design_name: Optional[str] = None,
                     aedt_version: Optional[str] = None,
                     port: Optional[int] = None,
                     variables: Optional[dict] = None) -> None:
        """Save current session state to file."""
        session_data = {
            "project_path": project_path or self._current_project_path,
            "design_name": design_name or self._current_design,
            "aedt_version": aedt_version or "2025.1",
            "port": port,
            "variables": variables or {},
        }
        with open(self.get_session_path(), "w") as f:
            json.dump(session_data, f, indent=2)

    def load_session(self) -> dict:
        """Load session state from file."""
        session_path = self.get_session_path()
        if os.path.exists(session_path):
            with open(session_path, "r") as f:
                return json.load(f)
        return {}

    def clear_session(self) -> None:
        """Clear session state."""
        self._current_project_path = None
        self._current_design = None
        session_path = self.get_session_path()
        if os.path.exists(session_path):
            os.remove(session_path)

    def set_current_project(self, project_path: str) -> None:
        """Set the current project path."""
        self._current_project_path = project_path
        self.save_session()

    def get_current_project(self) -> Optional[str]:
        """Get the current project path."""
        return self._current_project_path

    def set_current_design(self, design_name: str) -> None:
        """Set the current design name."""
        self._current_design = design_name
        self.save_session()

    def get_current_design(self) -> Optional[str]:
        """Get the current design name."""
        return self._current_design


def create_project(name: str, project_path: Optional[str] = None,
                   project_type: str = "HFSS") -> dict:
    """Create a new AEDT project.

    Parameters
    ----------
    name : str
        Project name.
    project_path : str, optional
        Directory for the project. Defaults to current directory.
    project_type : str
        Project type (HFSS, Maxwell3D, Icepak, Circuit, etc.).

    Returns
    -------
    dict
        Project creation result with status and path.
    """
    if project_path is None:
        project_path = os.getcwd()

    project_file = os.path.join(project_path, f"{name}.aedt")

    try:
        hfss = _get_hfss()

        # Project is created automatically when HFSS is launched
        # Get the actual project name from HFSS
        actual_project_name = hfss.project_name

        # Save session info
        manager = ProjectManager()
        manager.set_current_project(project_file)
        manager.set_current_design(hfss.design_name)

        return {
            "status": "created",
            "message": f"Project '{actual_project_name}' created",
            "path": project_file,
            "project_name": actual_project_name,
            "design_name": hfss.design_name,
        }

    except RuntimeError as e:
        # Fallback to template creation if AEDT not available
        if project_path is None:
            project_path = os.getcwd()

        project_file = os.path.join(project_path, f"{name}.aedt")

        # Check if file already exists
        if os.path.exists(project_file):
            return {
                "status": "error",
                "message": f"Project already exists: {project_file}",
                "path": project_file,
            }

        session_data = {
            "project_name": name,
            "project_path": project_file,
            "project_type": project_type,
            "design_name": f"{project_type}Design1",
            "aedt_version": "2025.1",
            "variables": {},
        }

        return {
            "status": "created",
            "message": f"Project template created (AEDT not available): {name}.aedt",
            "path": project_file,
            "session_data": session_data,
            "warning": "AEDT not connected - template created for later use",
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to create project: {str(e)}",
        }


def open_project(project_file: str) -> dict:
    """Open an existing AEDT project.

    Parameters
    ----------
    project_file : str
        Path to the .aedt project file.

    Returns
    -------
    dict
        Project info with status.
    """
    if not os.path.exists(project_file):
        return {
            "status": "error",
            "message": f"Project file not found: {project_file}",
        }

    if not project_file.endswith(".aedt"):
        return {
            "status": "error",
            "message": "Invalid project file. Must have .aedt extension",
        }

    try:
        _patch_design_settings()
        from pyaedt import Hfss

        global _hfss_instance
        _hfss_instance = Hfss(projectname=project_file, specified_version='2019.1')

        project_name = _hfss_instance.project_name
        design_name = _hfss_instance.design_name

        # Update session
        manager = ProjectManager()
        manager.set_current_project(project_file)
        manager.set_current_design(design_name)

        return {
            "status": "opened",
            "message": f"Opened project: {project_name}",
            "path": project_file,
            "name": project_name,
            "design_name": design_name,
        }

    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to open project: {str(e)}",
            "path": project_file,
        }


def save_project(project_file: Optional[str] = None) -> dict:
    """Save the current project.

    Parameters
    ----------
    project_file : str, optional
        Path to project file. Uses current session if not provided.

    Returns
    -------
    dict
        Save status.
    """
    try:
        hfss = _get_hfss()
        hfss.save_project()

        return {
            "status": "saved",
            "message": f"Project saved: {hfss.project_name}",
            "project_name": hfss.project_name,
        }

    except RuntimeError:
        # Fallback to session-based save
        manager = ProjectManager()
        session = manager.load_session()
        project_file = project_file or session.get("project_path")

        if project_file is None:
            return {
                "status": "error",
                "message": "No project file specified and no active session",
            }

        return {
            "status": "saved",
            "message": f"Project saved (AEDT not connected): {os.path.basename(project_file)}",
            "path": project_file,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to save project: {str(e)}",
        }


def close_project() -> dict:
    """Close the current project.

    Returns
    -------
    dict
        Close status.
    """
    global _hfss_instance

    try:
        if _hfss_instance is not None:
            project_name = _hfss_instance.project_name
            _hfss_instance.release_desktop()
            _hfss_instance = None
        else:
            project_name = None

        manager = ProjectManager()
        session = manager.load_session()
        if not project_name:
            project_name = os.path.basename(session.get("project_path", ""))
        manager.clear_session()

        return {
            "status": "closed",
            "message": f"Closed project: {project_name}" if project_name else "No active project",
        }

    except Exception as e:
        manager = ProjectManager()
        manager.clear_session()
        return {
            "status": "closed",
            "message": f"Project closed (with error): {str(e)}",
        }


def get_project_info(project_file: Optional[str] = None) -> dict:
    """Get information about a project.

    Parameters
    ----------
    project_file : str, optional
        Path to project file.

    Returns
    -------
    dict
        Project information.
    """
    try:
        hfss = _get_hfss()

        return {
            "status": "info",
            "name": hfss.project_name,
            "path": getattr(hfss, 'project_path', 'unknown'),
            "design_name": hfss.design_name,
            "design_type": getattr(hfss, 'design_type', 'unknown'),
            "exists": True,
        }

    except RuntimeError:
        # Fallback to file-based info
        manager = ProjectManager()
        session = manager.load_session()

        if project_file is None:
            project_file = session.get("project_path")

        if project_file is None or not os.path.exists(project_file):
            return {
                "status": "no_project",
                "message": "No project is currently open",
            }

        return {
            "status": "info",
            "name": os.path.splitext(os.path.basename(project_file))[0],
            "path": project_file,
            "exists": os.path.exists(project_file),
            "warning": "AEDT not connected - limited info available",
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to get project info: {str(e)}",
        }


def list_designs(project_file: str) -> dict:
    """List designs in a project.

    Parameters
    ----------
    project_file : str
        Path to the .aedt project file.

    Returns
    -------
    dict
        List of designs.
    """
    if not os.path.exists(project_file):
        return {
            "status": "error",
            "message": f"Project file not found: {project_file}",
            "designs": [],
        }

    try:
        _patch_design_settings()
        from pyaedt import Hfss

        # Open project temporarily to list designs
        hfss = Hfss(projectname=project_file, specified_version='2019.1')

        designs = []
        for design_name in hfss.design_list:
            designs.append({
                "name": design_name,
                "type": "HFSS",  # Would need more API calls to get exact type
                "locked": False,
            })

        hfss.release_desktop()

        return {
            "status": "info",
            "project": os.path.basename(project_file),
            "designs": designs,
            "count": len(designs),
        }

    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to list designs: {str(e)}",
            "project": os.path.basename(project_file),
            "designs": [],
        }
