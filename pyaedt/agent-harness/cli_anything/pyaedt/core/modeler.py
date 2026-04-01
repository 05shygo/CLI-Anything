"""PyAEDT modeler operations - Core module."""

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


def create_box(name: str, position: list, dimensions: list,
               material: str = "aluminum") -> dict:
    """Create a box in the modeler.

    Parameters
    ----------
    name : str
        Name for the box.
    position : list
        Position [x, y, z] with units (e.g., [0, 0, 0]).
    dimensions : list
        Dimensions [dx, dy, dz] with units (e.g., [1, 1, 1]).
    material : str
        Material name.

    Returns
    -------
    dict
        Creation result.
    """
    if len(position) != 3:
        return {
            "status": "error",
            "message": "Position must have 3 coordinates [x, y, z]",
        }

    if len(dimensions) != 3:
        return {
            "status": "error",
            "message": "Dimensions must have 3 values [dx, dy, dz]",
        }

    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        box_obj = app.modeler.create_box(
            origin=position,
            sizes=dimensions,
            name=name,
            material=material,
        )
        if box_obj:
            return {
                "status": "created",
                "message": f"Box '{name}' created at {position} with dimensions {dimensions}",
                "object_name": name,
                "type": "Box",
                "position": position,
                "dimensions": dimensions,
                "material": material,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to create box '{name}'",
            }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }


def create_cylinder(name: str, position: list, radius: float,
                    height: float, material: str = "aluminum",
                    axis: str = "Z") -> dict:
    """Create a cylinder in the modeler.

    Parameters
    ----------
    name : str
        Name for the cylinder.
    position : list
        Center position [x, y, z].
    radius : float
        Radius with units (e.g., 1).
    height : float
        Height with units (e.g., 10).
    material : str
        Material name.
    axis : str
        Axis of extrusion: "X", "Y", or "Z".

    Returns
    -------
    dict
        Creation result.
    """
    if axis not in ["X", "Y", "Z"]:
        return {
            "status": "error",
            "message": f"Invalid axis: {axis}. Must be X, Y, or Z",
        }

    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        cyl_obj = app.modeler.create_cylinder(
            orientation=axis,
            origin=position,
            radius=radius,
            height=height,
            name=name,
            material=material,
        )
        if cyl_obj:
            return {
                "status": "created",
                "message": f"Cylinder '{name}' created at {position}",
                "object_name": name,
                "type": "Cylinder",
                "position": position,
                "radius": radius,
                "height": height,
                "axis": axis,
                "material": material,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to create cylinder '{name}'",
            }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }


def create_sphere(name: str, position: list, radius: float,
                 material: str = "aluminum") -> dict:
    """Create a sphere in the modeler.

    Parameters
    ----------
    name : str
        Name for the sphere.
    position : list
        Center position [x, y, z].
    radius : float
        Radius with units.
    material : str
        Material name.

    Returns
    -------
    dict
        Creation result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        sphere_obj = app.modeler.create_sphere(
            origin=position,
            radius=radius,
            name=name,
            material=material,
        )
        if sphere_obj:
            return {
                "status": "created",
                "message": f"Sphere '{name}' created at {position}",
                "object_name": name,
                "type": "Sphere",
                "position": position,
                "radius": radius,
                "material": material,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to create sphere '{name}'",
            }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }


def assign_material(object_name: str, material: str) -> dict:
    """Assign material to an object.

    Parameters
    ----------
    object_name : str
        Name of the object.
    material : str
        Material name.

    Returns
    -------
    dict
        Assignment result.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        result = app.assign_material(assignment=object_name, material=material)
        if result:
            return {
                "status": "assigned",
                "message": f"Material '{material}' assigned to '{object_name}'",
                "object_name": object_name,
                "material": material,
            }
        else:
            return {
                "status": "error",
                "message": f"Failed to assign material '{material}' to '{object_name}'",
            }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }


def list_objects() -> dict:
    """List all objects in the modeler.

    Returns
    -------
    dict
        List of objects.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "objects": [],
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        objects = list(app.modeler.objects_by_name.keys())
        return {
            "status": "info",
            "objects": objects,
            "message": f"Found {len(objects)} objects",
        }
    except Exception as e:
        return {
            "status": "error",
            "objects": [],
            "message": str(e),
        }


def delete_object(object_name: str) -> dict:
    """Delete an object from the modeler.

    Parameters
    ----------
    object_name : str
        Name of the object to delete.

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
        obj = app.modeler.objects_by_name.get(object_name)
        if obj is None:
            return {
                "status": "error",
                "message": f"Object '{object_name}' not found",
            }
        obj.delete()
        return {
            "status": "deleted",
            "message": f"Object '{object_name}' deleted",
            "object_name": object_name,
        }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }


def get_object_info(object_name: str) -> dict:
    """Get information about an object.

    Parameters
    ----------
    object_name : str
        Name of the object.

    Returns
    -------
    dict
        Object information.
    """
    app = _get_aedt_app()
    if app is None:
        return {
            "status": "error",
            "message": "AEDT is not available. Connect to an AEDT session first.",
        }

    try:
        obj = app.modeler.objects_by_name.get(object_name)
        if obj is None:
            return {
                "status": "error",
                "message": f"Object '{object_name}' not found",
            }
        return {
            "status": "info",
            "object_name": obj.name,
            "type": obj.object_type,
            "volume": obj.volume,
            "surface_area": obj.surface_area,
            "material": obj.material,
            "message": f"Info for object '{object_name}'",
        }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }
