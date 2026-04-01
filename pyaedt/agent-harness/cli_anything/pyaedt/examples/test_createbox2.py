"""Minimal test: create one box in AEDT 2019.1 via COM. Output to file."""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

logfile = r"D:\class_design\test_output.log"
class Logger:
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")
    def write(self, msg):
        self.terminal.write(msg)
        self.log.write(msg)
        self.log.flush()
    def flush(self):
        self.terminal.flush()
        self.log.flush()
sys.stdout = Logger(logfile)
sys.stderr = sys.stdout

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
    projectname=r"D:\class_design\test_box2",
    designname="TestDesign",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
print(f"Project: {hfss.project_name}")
print(f"Design: {hfss.design_name}")
hfss.modeler.model_units = "mm"

# Method 1: PyAEDT high-level API
print("\n--- Method 1: hfss.modeler.create_box (high-level API) ---")
try:
    box = hfss.modeler.create_box(
        origin=[-10, -10, 0],
        sizes=[20, 20, 1],
        name="TestBox_HL",
        material="FR4_epoxy",
    )
    print(f"  Success: {box}")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()

# Method 2: COM with oeditor
print("\n--- Method 2: COM oeditor ---")
try:
    od = hfss.modeler.oeditor
    print(f"  oeditor type: {type(od)}")
    print(f"  oeditor: {od}")
except Exception as e:
    print(f"  Failed to get oeditor: {e}")
    traceback.print_exc()
    od = None

# Method 3: Try different editor access
print("\n--- Method 3: hfss.odesign.SetActiveEditor('3D Modeler') ---")
try:
    od2 = hfss.odesign.SetActiveEditor("3D Modeler")
    print(f"  od2 type: {type(od2)}")
    print(f"  od2: {od2}")
except Exception as e:
    print(f"  Failed: {e}")
    traceback.print_exc()
    od2 = None

if od2:
    print("\n--- Method 4: COM CreateBox via SetActiveEditor ---")
    try:
        od2.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "-60",
                "YPosition:=", "-30",
                "ZPosition:=", "0",
                "XSize:=", "120",
                "YSize:=", "60",
                "ZSize:=", "1.6"
            ],
            [
                "NAME:Attributes",
                "Name:=", "TestBox_COM",
                "Flags:=", "",
                "Color:=", "(143 175 131)",
                "Transparency:=", 0,
                "PartCoordinateSystem:=", "Global",
                "UDMId:=", "",
                "MaterialValue:=", '"FR4_epoxy"',
                "SurfaceMaterialValue:=", '""',
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False,
                "IsLightweight:=", False
            ]
        )
        print("  Success!")
    except Exception as e:
        print(f"  Failed: {e}")
        traceback.print_exc()

    # Try with XSizing instead of XSize
    print("\n--- Method 5: COM CreateBox with XSizing ---")
    try:
        od2.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "-60",
                "YPosition:=", "-30",
                "ZPosition:=", "2",
                "XSizing:=", "120",
                "YSizing:=", "60",
                "ZSizing:=", "1.6"
            ],
            [
                "NAME:Attributes",
                "Name:=", "TestBox_COM2",
                "Flags:=", "",
                "Color:=", "(143 175 131)",
                "Transparency:=", 0,
                "PartCoordinateSystem:=", "Global",
                "UDMId:=", "",
                "MaterialValue:=", '"FR4_epoxy"',
                "SurfaceMaterialValue:=", '""',
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False,
                "IsLightweight:=", False
            ]
        )
        print("  Success!")
    except Exception as e:
        print(f"  Failed: {e}")
        traceback.print_exc()

print("\n\nAll tests done.")
try:
    hfss.release_desktop()
except:
    pass
