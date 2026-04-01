"""PyAEDT session management - Core module."""

from __future__ import annotations

import os
import sys
from typing import Optional


class SessionManager:
    """Manages AEDT session state and connections."""

    def __init__(self):
        """Initialize session manager."""
        self._port: Optional[int] = None
        self._process_id: Optional[int] = None
        self._version: Optional[str] = None
        self._connected: bool = False

    @property
    def port(self) -> Optional[int]:
        """Get the current gRPC port."""
        return self._port

    @property
    def process_id(self) -> Optional[int]:
        """Get the AEDT process ID."""
        return self._process_id

    @property
    def version(self) -> Optional[str]:
        """Get the AEDT version."""
        return self._version

    @property
    def is_connected(self) -> bool:
        """Check if connected to AEDT."""
        return self._connected

    def set_connection(self, port: int, process_id: int,
                      version: str) -> None:
        """Set connection parameters."""
        self._port = port
        self._process_id = process_id
        self._version = version
        self._connected = True

    def disconnect(self) -> None:
        """Disconnect from AEDT session."""
        self._port = None
        self._process_id = None
        self._version = None
        self._connected = False

    def get_status(self) -> dict:
        """Get session status."""
        if self._connected:
            return {
                "status": "connected",
                "port": self._port,
                "process_id": self._process_id,
                "version": self._version,
            }
        else:
            return {
                "status": "disconnected",
                "message": "Not connected to AEDT",
            }


def start_session(version: str = "2025.1", non_graphical: bool = False,
                  port: int = 0) -> dict:
    """Start a new AEDT session.

    Parameters
    ----------
    version : str
        AEDT version (e.g., "2025.1").
    non_graphical : bool
        Run in non-graphical mode.
    port : int
        gRPC port (0 for auto).

    Returns
    -------
    dict
        Session start result.
    """
    # Check if running on Windows
    if sys.platform != "win32":
        return {
            "status": "error",
            "message": "AEDT is only available on Windows",
            "platform": sys.platform,
        }

    return {
        "status": "starting",
        "message": f"Starting AEDT {version}...",
        "version": version,
        "non_graphical": non_graphical,
        "port": port,
    }


def attach_session(port: int) -> dict:
    """Attach to an existing AEDT session.

    Parameters
    ----------
    port : int
        gRPC port of the AEDT session.

    Returns
    -------
    dict
        Attach result.
    """
    return {
        "status": "attached",
        "message": f"Attached to AEDT on port {port}",
        "port": port,
    }


def detach_session() -> dict:
    """Detach from the current AEDT session.

    Returns
    -------
    dict
        Detach result.
    """
    return {
        "status": "detached",
        "message": "Detached from AEDT session",
    }


def get_session_status() -> dict:
    """Get the current session status.

    Returns
    -------
    dict
        Session status.
    """
    return {
        "status": "info",
        "connected": False,
        "message": "Connect to AEDT to check session status",
    }


def is_aedt_available() -> dict:
    """Check if AEDT is installed and available.

    Returns
    -------
    dict
        Availability status.
    """
    import shutil

    # Check for ansysedt in PATH
    aedt_path = shutil.which("ansysedt")
    if aedt_path:
        return {
            "status": "available",
            "path": aedt_path,
            "message": "AEDT is installed",
        }

    return {
        "status": "unavailable",
        "message": "AEDT not found in PATH. Install Ansys Electronics Desktop.",
    }
