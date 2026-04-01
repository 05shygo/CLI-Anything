"""Minimal test: create one box in AEDT 2019.1 via COM."""
import os
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

# Patch for 2019.1
try:
    from pyaedt import desktop as _dm
    _orig = _dm.Desktop.__init__
    def _p(self, *a, **kw):
        _orig(self, *a, **kw)
        if not hasattr(self, 'student_version'):
            self.student_version = False
    _dm.Desktop.__init__ = _p
except Exception:
    pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(self, app):
        try: _o2(self, app)
        except AttributeError:
            self._app = app; self.design_settings = None; self.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except Exception:
    pass

from pyaedt import Hfss
import traceback

print("Launching HFSS...")
hfss = Hfss(
    projectname=r"D:\class_design\test_box",
    designname="TestDesign",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
print("Project:", hfss.project_name)
print("Design:", hfss.design_name)

hfss.modeler.model_units = "mm"

# Method 1: PyAEDT high-level API
print("\n--- Method 1: hfss.modeler.create_box ---")
try:
    box = hfss.modeler.create_box(
        origin=[-10, -10, 0],
        sizes=[20, 20, 1],
        name="TestBox1",
        material="FR4_epoxy",
    )
    print(f"  Success: {box}")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()

# Method 2: COM direct - original format from working example
print("\n--- Method 2: COM CreateBox (string vals) ---")
try:
    oEditor = hfss.modeler.oeditor
    print(f"  oEditor type: {type(oEditor)}")
    oEditor.CreateBox(
        [
            "NAME:BoxParameters",
            "XPosition:=", "-20",
            "YPosition:=", "-10",
            "ZPosition:=", "0",
            "XSizing:=", "40",
            "YSizing:=", "20",
            "ZSizing:=", "1.6"
        ],
        [
            "NAME:Attributes",
            "Name:=", "TestBox2",
            "Flags:=", "",
            "Color:=", "(143 175 131)",
            "Transparency:=", 0,
            "PartCoordinateSystem:=", "Global",
            "UDMId:=", "",
            "MaterialValue:=", '"FR4_epoxy"',
            "SurfaceMaterialValue:=", '""',
            "SolveInside:=", "true",
            "IsMaterialEditable:=", True,
            "UseMaterialAppearance:=", False,
            "IsLightweight:=", False
        ]
    )
    print("  Success!")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()

# Method 3: COM with "mm" units
print("\n--- Method 3: COM CreateBox (with mm units) ---")
try:
    oEditor.CreateBox(
        [
            "NAME:BoxParameters",
            "XPosition:=", "-30mm",
            "YPosition:=", "-10mm",
            "ZPosition:=", "0mm",
            "XSizing:=", "60mm",
            "YSizing:=", "20mm",
            "ZSizing:=", "1.6mm"
        ],
        [
            "NAME:Attributes",
            "Name:=", "TestBox3",
            "Flags:=", "",
            "Color:=", "(143 175 131)",
            "Transparency:=", 0,
            "PartCoordinateSystem:=", "Global",
            "UDMId:=", "",
            "MaterialValue:=", '"FR4_epoxy"',
            "SurfaceMaterialValue:=", '""',
            "SolveInside:=", "true",
            "IsMaterialEditable:=", True,
            "UseMaterialAppearance:=", False,
            "IsLightweight:=", False
        ]
    )
    print("  Success!")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()

# Method 4: Original example  with UDMList not UDMId
print("\n--- Method 4: COM CreateBox (UDMList) ---")
try:
    oEditor.CreateBox(
        [
            "NAME:BoxParameters",
            "XPosition:=", "-40",
            "YPosition:=", "-10",
            "ZPosition:=", "0",
            "XSizing:=", "80",
            "YSizing:=", "20",
            "ZSizing:=", "1.6"
        ],
        [
            "NAME:Attributes",
            "Name:=", "TestBox4",
            "Flags:=", "",
            "Color:=", "(143 175 131)",
            "Transparency:=", 0,
            "PartCoordinateSystem:=", "Global",
            "UDMList:=", [],
            "MaterialValue:=", '"FR4_epoxy"',
            "SurfaceMaterialValue:=", '""',
            "SolveInside:=", "true",
            "IsMaterialEditable:=", True,
            "UseMaterialAppearance:=", False,
            "IsLightweight:=", False
        ]
    )
    print("  Success!")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()

print("\nAll tests done.")
hfss.release_desktop()
