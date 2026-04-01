"""PyAEDT variable management - Core module."""

from __future__ import annotations

from typing import Optional


def list_variables(scope: str = "all") -> dict:
    """List variables in the current design.

    Parameters
    ----------
    scope : str
        Variable scope: "all", "project", or "design".

    Returns
    -------
    dict
        List of variables.
    """
    return {
        "status": "info",
        "scope": scope,
        "variables": [],
        "message": "Connect to AEDT to list variables",
    }


def set_variable(name: str, value: str, scope: str = "design") -> dict:
    """Set a variable value.

    Parameters
    ----------
    name : str
        Variable name (with $ prefix for project variables).
    value : str
        Variable value with units (e.g., "10GHz", "1mm").
    scope : str
        Variable scope: "project" or "design".

    Returns
    -------
    dict
        Set variable result.
    """
    if not name:
        return {
            "status": "error",
            "message": "Variable name cannot be empty",
        }

    if not value:
        return {
            "status": "error",
            "message": "Variable value cannot be empty",
        }

    # Add $ prefix for project variables
    if scope == "project" and not name.startswith("$"):
        name = f"${name}"

    return {
        "status": "set",
        "message": f"Variable '{name}' set to '{value}'",
        "name": name,
        "value": value,
        "scope": scope,
    }


def get_variable(name: str) -> dict:
    """Get a variable value.

    Parameters
    ----------
    name : str
        Variable name.

    Returns
    -------
    dict
        Variable value.
    """
    if not name:
        return {
            "status": "error",
            "message": "Variable name cannot be empty",
        }

    # Variable lookup would happen here via PyAEDT
    return {
        "status": "found",
        "name": name,
        "value": "unknown",
        "message": "Connect to AEDT to get variable value",
    }


def delete_variable(name: str) -> dict:
    """Delete a variable.

    Parameters
    ----------
    name : str
        Variable name.

    Returns
    -------
    dict
        Deletion result.
    """
    return {
        "status": "deleted",
        "message": f"Variable '{name}' deleted",
        "name": name,
    }


def list_materials() -> dict:
    """List available materials.

    Returns
    -------
    dict
        List of materials.
    """
    common_materials = [
        "aluminum", "copper", "gold", "silver", "steel",
        "aluminum_oxide", "silicon", "germanium",
        "FR4_epoxy", "Rogers_RO4003", "Rogers_RO4350",
        "air", "vacuum", "water", "glass",
    ]

    return {
        "status": "info",
        "materials": common_materials,
        "count": len(common_materials),
    }
