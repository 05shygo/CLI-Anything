"""PyAEDT export operations - Core module."""

from __future__ import annotations

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


def export_results(output_path: str, format: str = "csv",
                  setup_name: Optional[str] = None) -> dict:
    """Export simulation results.

    Parameters
    ----------
    output_path : str
        Path for the output file.
    format : str
        Export format: "csv", "hdf5", "snp", "jpg", "png", "pdf".
    setup_name : str, optional
        Name of the setup to export from.

    Returns
    -------
    dict
        Export result.
    """
    supported_formats = ["csv", "hdf5", "snp", "jpg", "png", "pdf", "step", "stl"]

    if format.lower() not in supported_formats:
        return {
            "status": "error",
            "message": f"Unsupported format: {format}",
            "supported_formats": supported_formats,
        }

    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        if format.lower() == "csv":
            # Export results using the analysis export_results method
            exported = app.export_results(export_folder=output_path)
            return {
                "status": "exported",
                "message": f"Results exported to {output_path}",
                "path": output_path,
                "format": format,
                "setup_name": setup_name,
            }
        elif format.lower() == "snp":
            # Export touchstone file
            app.export_touchstone(
                setup=setup_name,
                output_file=output_path,
            )
            return {
                "status": "exported",
                "message": f"Touchstone file exported to {output_path}",
                "path": output_path,
                "format": format,
                "setup_name": setup_name,
            }
        elif format.lower() in ["jpg", "png"]:
            # Export model image
            app.post.plot_model_obj(
                show=False,
                export_path=output_path,
            )
            return {
                "status": "exported",
                "message": f"Image exported to {output_path}",
                "path": output_path,
                "format": format,
            }
        elif format.lower() == "step":
            app.modeler.export_step(output_path)
            return {
                "status": "exported",
                "message": f"STEP exported to {output_path}",
                "path": output_path,
                "format": format,
            }
        elif format.lower() == "stl":
            app.modeler.export_stl(output_path)
            return {
                "status": "exported",
                "message": f"STL exported to {output_path}",
                "path": output_path,
                "format": format,
            }
        else:
            return {
                "status": "error",
                "message": f"Unsupported format: {format}",
                "supported_formats": supported_formats,
            }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Export failed: {str(e)}",
            "format": format,
        }


def export_image(output_path: str, width: int = 1920,
                height: int = 1080) -> dict:
    """Export project image.

    Parameters
    ----------
    output_path : str
        Path for the output image.
    width : int
        Image width in pixels.
    height : int
        Image height in pixels.

    Returns
    -------
    dict
        Export result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        app.post.plot_model_obj(
            show=False,
            export_path=output_path,
        )
        return {
            "status": "exported",
            "message": f"Image exported to {output_path}",
            "path": output_path,
            "width": width,
            "height": height,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Image export failed: {str(e)}",
        }


def export_to_hfss(output_path: str) -> dict:
    """Export project to HFSS format.

    This saves the current project to an HFSS project file.

    Parameters
    ----------
    output_path : str
        Path for the HFSS file.

    Returns
    -------
    dict
        Export result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        # Save project to the specified path
        app.save_project(project_path=output_path)
        return {
            "status": "exported",
            "message": f"Exported to HFSS: {output_path}",
            "path": output_path,
            "format": "hfss",
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"HFSS export failed: {str(e)}",
        }


def export_to_step(output_path: str) -> dict:
    """Export geometry to STEP format.

    Parameters
    ----------
    output_path : str
        Path for the STEP file.

    Returns
    -------
    dict
        Export result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        app.modeler.export_step(output_path)
        return {
            "status": "exported",
            "message": f"Exported to STEP: {output_path}",
            "path": output_path,
            "format": "step",
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"STEP export failed: {str(e)}",
        }


def export_to_stl(output_path: str, scale: float = 1.0) -> dict:
    """Export geometry to STL format.

    Parameters
    ----------
    output_path : str
        Path for the STL file.
    scale : float
        Scale factor.

    Returns
    -------
    dict
        Export result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        app.modeler.export_stl(output_path, scale=scale)
        return {
            "status": "exported",
            "message": f"Exported to STL: {output_path}",
            "path": output_path,
            "format": "stl",
            "scale": scale,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"STL export failed: {str(e)}",
        }


def get_export_formats() -> dict:
    """Get list of supported export formats.

    Returns
    -------
    dict
        List of export formats.
    """
    return {
        "status": "info",
        "formats": {
            "results": ["csv", "hdf5", "snp", "mat"],
            "images": ["jpg", "png", "bmp", "gif"],
            "geometry": ["step", "stl", "iges", "sat"],
            "documents": ["pdf", "docx"],
        },
    }
