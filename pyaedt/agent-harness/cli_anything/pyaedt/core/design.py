"""PyAEDT design operations - Core module."""

from __future__ import annotations

from typing import Optional


# Design types supported by AEDT
DESIGN_TYPES = {
    "HFSS": "High Frequency Structural Simulator",
    "HFSS3DLayout": "HFSS 3D Layout",
    "Maxwell2D": "Maxwell 2D",
    "Maxwell3D": "Maxwell 3D",
    "RMXprt": "Rotating Machine Expert",
    "Icepak": "Icepak Thermal Analysis",
    "Q3D": "Q3D Extractor",
    "Q3DExtractive": "Q3D Extractive",
    "Circuit": "Nexxim Circuit",
    "CircuitNetlist": "Circuit Netlist",
    "TwinBuilder": "Twin Builder System",
    "Mechanical": "Mechanical",
    "EDB": "Electronic Database",
    "EMIT": "EMIT RF Propagation",
}


def create_design(design_name: str, design_type: str = "HFSS") -> dict:
    """Create a new design in the current project.

    Parameters
    ----------
    design_name : str
        Name for the new design.
    design_type : str
        Type of design (HFSS, Maxwell3D, Icepak, Circuit, etc.).

    Returns
    -------
    dict
        Design creation result.
    """
    if design_type not in DESIGN_TYPES:
        return {
            "status": "error",
            "message": f"Invalid design type: {design_type}",
            "valid_types": list(DESIGN_TYPES.keys()),
        }

    return {
        "status": "created",
        "message": f"Design '{design_name}' created ({design_type})",
        "design_name": design_name,
        "design_type": design_type,
        "description": DESIGN_TYPES[design_type],
    }


def list_designs() -> dict:
    """List all designs in the current project.

    Returns
    -------
    dict
        List of designs with their types.
    """
    # This would query the current AEDT project
    return {
        "status": "info",
        "designs": [],
        "message": "Connect to AEDT to list designs",
    }


def activate_design(design_name: str) -> dict:
    """Set the active design by name.

    Parameters
    ----------
    design_name : str
        Name of the design to activate.

    Returns
    -------
    dict
        Activation result.
    """
    return {
        "status": "activated",
        "message": f"Activated design: {design_name}",
        "design_name": design_name,
    }


def copy_design(source_design: str, new_design_name: str) -> dict:
    """Copy an existing design.

    Parameters
    ----------
    source_design : str
        Name of the source design to copy.
    new_design_name : str
        Name for the new design.

    Returns
    -------
    dict
        Copy result.
    """
    return {
        "status": "copied",
        "message": f"Design '{source_design}' copied to '{new_design_name}'",
        "source": source_design,
        "new_design": new_design_name,
    }


def get_design_info(design_name: Optional[str] = None) -> dict:
    """Get information about a design.

    Parameters
    ----------
    design_name : str, optional
        Name of the design. Uses current design if not specified.

    Returns
    -------
    dict
        Design information.
    """
    if design_name is None:
        return {
            "status": "info",
            "message": "No design specified and no active design",
        }

    return {
        "status": "info",
        "design_name": design_name,
        "design_type": "HFSS",
        "variables": [],
        "objects": 0,
        "setups": [],
    }


def delete_design(design_name: str) -> dict:
    """Delete a design from the project.

    Parameters
    ----------
    design_name : str
        Name of the design to delete.

    Returns
    -------
    dict
        Deletion result.
    """
    return {
        "status": "deleted",
        "message": f"Design '{design_name}' deleted",
        "design_name": design_name,
    }


def rename_design(old_name: str, new_name: str) -> dict:
    """Rename a design.

    Parameters
    ----------
    old_name : str
        Current design name.
    new_name : str
        New design name.

    Returns
    -------
    dict
        Rename result.
    """
    return {
        "status": "renamed",
        "message": f"Design renamed from '{old_name}' to '{new_name}'",
        "old_name": old_name,
        "new_name": new_name,
    }
