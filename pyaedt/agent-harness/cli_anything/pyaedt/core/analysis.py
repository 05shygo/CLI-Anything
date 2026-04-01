"""PyAEDT analysis operations - Core module."""

from __future__ import annotations

import re
from typing import Optional

# Lazy-import PyAEDT to allow the module to load without AEDT installed
_aedt_app = None


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


def _get_aedt_app():
    """Get or create the global AEDT application instance."""
    global _aedt_app
    if _aedt_app is not None:
        return _aedt_app
    try:
        _patch_design_settings()
        from pyaedt import Hfss
        _aedt_app = Hfss(non_graphical=True, specified_version='2019.1')
        return _aedt_app
    except Exception:
        return None


def _parse_frequency(value):
    """Parse a frequency value string like '10GHz' or float like 10e9.

    Returns tuple of (value, unit). Defaults to GHz if no unit.
    """
    if isinstance(value, (int, float)):
        return float(value), "GHz"
    # String like "10GHz", "1.5MHz", etc.
    match = re.match(r'^([0-9.]+)\s*(GHz|MHz|kHz|Hz)?$', str(value), re.IGNORECASE)
    if match:
        return float(match.group(1)), match.group(2) or "GHz"
    # Fallback: treat as-is with GHz
    try:
        return float(value), "GHz"
    except ValueError:
        return float(value), "GHz"


def create_setup(setup_name: str, setup_type: str = "Hfss",
                properties: Optional[dict] = None) -> dict:
    """Create an analysis setup.

    Parameters
    ----------
    setup_name : str
        Name for the setup.
    setup_type : str
        Type of setup: "Hfss", "Maxwell", "Icepak", "Circuit", etc.
    properties : dict, optional
        Setup properties (frequency, sweeps, etc.).

    Returns
    -------
    dict
        Setup creation result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    default_props = {
        "Hfss": {
            "frequency": "10GHz",
            "max_delta_z": 0.02,
            "max_passes": 20,
        },
        "Maxwell": {
            "frequency": "60Hz",
            "max_passes": 20,
        },
        "Icepak": {
            "max_iterations": 100,
        },
    }

    props = properties or default_props.get(setup_type, {})

    try:
        setup = app.create_setup(
            name=setup_name,
            setup_type=setup_type,
            Frequency=props.get("frequency", "10GHz"),
            MaxDeltaZ=props.get("max_delta_z", 0.02),
            MaxPasses=props.get("max_passes", 20),
        )
        if setup:
            return {
                "status": "created",
                "message": f"Analysis setup '{setup_name}' created ({setup_type})",
                "setup_name": setup_name,
                "setup_type": setup_type,
                "properties": props,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to create setup '{setup_name}'",
                "setup_name": setup_name,
                "setup_type": setup_type,
            }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to create setup '{setup_name}': {str(e)}",
            "setup_name": setup_name,
            "setup_type": setup_type,
        }


def list_setups() -> dict:
    """List all analysis setups.

    Returns
    -------
    dict
        List of setups.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "setups": [],
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        setups = app.get_setups()
        return {
            "status": "info",
            "setups": setups,
            "message": f"Found {len(setups)} setups",
        }
    except Exception as e:
        return {
            "status": "error",
            "setups": [],
            "message": str(e),
        }


def run_setup(setup_name: str) -> dict:
    """Run an analysis setup.

    Parameters
    ----------
    setup_name : str
        Name of the setup to run.

    Returns
    -------
    dict
        Run result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        result = app.analyze_setup(setup_name)
        if result:
            return {
                "status": "running",
                "message": f"Analysis setup '{setup_name}' started",
                "setup_name": setup_name,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to run setup '{setup_name}'",
                "setup_name": setup_name,
            }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to run setup '{setup_name}': {str(e)}",
            "setup_name": setup_name,
        }


def delete_setup(setup_name: str) -> dict:
    """Delete an analysis setup.

    Parameters
    ----------
    setup_name : str
        Name of the setup to delete.

    Returns
    -------
    dict
        Deletion result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        result = app.delete_setup(setup_name)
        if result:
            return {
                "status": "deleted",
                "message": f"Analysis setup '{setup_name}' deleted",
                "setup_name": setup_name,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to delete setup '{setup_name}'",
                "setup_name": setup_name,
            }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to delete setup '{setup_name}': {str(e)}",
            "setup_name": setup_name,
        }


def get_setup_info(setup_name: str) -> dict:
    """Get information about an analysis setup.

    Parameters
    ----------
    setup_name : str
        Name of the setup.

    Returns
    -------
    dict
        Setup information.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        setups = app.get_setups()
        if setup_name in setups:
            return {
                "status": "info",
                "setup_name": setup_name,
                "type": "Hfss",
                "properties": {},
                "message": f"Setup '{setup_name}' found",
            }
        else:
            return {
                "status": "error",
                "setup_name": setup_name,
                "type": None,
                "properties": {},
                "message": f"Setup '{setup_name}' not found",
            }
    except Exception as e:
        return {
            "status": "error",
            "setup_name": setup_name,
            "type": None,
            "properties": {},
            "message": str(e),
        }


def add_frequency_sweep(setup_name: str, sweep_name: str,
                       start: float, end: float,
                       step: Optional[float] = None,
                       sweep_type: str = "Linear") -> dict:
    """Add a frequency sweep to a setup.

    Parameters
    ----------
    setup_name : str
        Name of the setup.
    sweep_name : str
        Name for the sweep.
    start : float
        Start frequency with units (e.g., "1GHz" or 1e9).
    end : float
        End frequency with units (e.g., "10GHz" or 10e9).
    step : float, optional
        Step size or count depending on sweep_type.
    sweep_type : str
        Sweep type: "Linear" (count-based), "Logarithmic", "Points".

    Returns
    -------
    dict
        Sweep creation result.
    """
    if sweep_type not in ["Linear", "Logarithmic", "Points"]:
        return {
            "status": "error",
            "message": f"Invalid sweep type: {sweep_type}",
        }

    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        start_val, unit = _parse_frequency(start)
        end_val, _ = _parse_frequency(end)

        if sweep_type == "Linear":
            # Use count-based sweep if step is provided as count
            if step is not None:
                # Try to interpret step - if it's a small number, treat as count
                # if it's a frequency value, treat as step size
                if step < 1000:
                    # Likely a count
                    sweep = app.create_linear_count_sweep(
                        setup=setup_name,
                        unit=unit,
                        start_frequency=start_val,
                        stop_frequency=end_val,
                        num_of_freq_points=int(step),
                        name=sweep_name,
                    )
                else:
                    # Treat as step size frequency
                    sweep = app.create_linear_step_sweep(
                        setup=setup_name,
                        unit=unit,
                        start_frequency=start_val,
                        stop_frequency=end_val,
                        step_size=step,
                        name=sweep_name,
                    )
            else:
                # Default count sweep with 1001 points
                sweep = app.create_linear_count_sweep(
                    setup=setup_name,
                    unit=unit,
                    start_frequency=start_val,
                    stop_frequency=end_val,
                    num_of_freq_points=1001,
                    name=sweep_name,
                )
        elif sweep_type == "Points":
            # Single point sweep at each frequency point
            if step is not None and step < 1000:
                # Multiple single point sweeps
                sweep = app.create_single_point_sweep(
                    setup=setup_name,
                    unit=unit,
                    freq=[start_val + i * (end_val - start_val) / (step - 1)
                          for i in range(int(step))] if step > 1 else [start_val],
                    name=sweep_name,
                )
            else:
                sweep = app.create_single_point_sweep(
                    setup=setup_name,
                    unit=unit,
                    freq=start_val,
                    name=sweep_name,
                )
        else:
            # Logarithmic - approximate with count sweep
            sweep = app.create_linear_count_sweep(
                setup=setup_name,
                unit=unit,
                start_frequency=start_val,
                stop_frequency=end_val,
                num_of_freq_points=int(step) if step else 1001,
                name=sweep_name,
                sweep_type="Interpolating",
            )

        if sweep:
            return {
                "status": "created",
                "message": f"Sweep '{sweep_name}' added to '{setup_name}'",
                "setup_name": setup_name,
                "sweep_name": sweep_name,
                "start": start,
                "end": end,
                "step": step,
                "sweep_type": sweep_type,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to add sweep to '{setup_name}'",
                "setup_name": setup_name,
                "sweep_name": sweep_name,
            }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to add sweep to '{setup_name}': {str(e)}",
            "setup_name": setup_name,
            "sweep_name": sweep_name,
        }
