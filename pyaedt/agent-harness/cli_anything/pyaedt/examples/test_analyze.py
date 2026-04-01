"""Test: open existing project and try analyze with validation check."""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "analyze_test.log")

class DL:
    def __init__(self, p):
        self.t = sys.stdout
        self.f = open(p, "w", encoding="utf-8")
    def write(self, m):
        try: self.t.write(m)
        except: pass
        self.f.write(m); self.f.flush()
    def flush(self):
        self.t.flush(); self.f.flush()

sys.stdout = DL(LOG)
sys.stderr = sys.stdout

try:
    from pyaedt import desktop as _dm
    _o = _dm.Desktop.__init__
    def _p(s, *a, **k):
        _o(s, *a, **k)
        if not hasattr(s, 'student_version'): s.student_version = False
    _dm.Desktop.__init__ = _p
except: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(s, app):
        try: _o2(s, app)
        except AttributeError:
            s._app = app; s.design_settings = None; s.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except: pass

from pyaedt import Hfss
import traceback

print("Opening existing project Printed_Dipole_v5...")
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Printed_Dipole_v5"),
    designname="Dipole_Balun",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
print(f"Project: {hfss.project_name}")
print(f"Design:  {hfss.design_name}")

oDesign = hfss.odesign
oDesktop = hfss.odesktop

# Check validation
print("\n--- Check Design Validation ---")
try:
    # Try to validate
    result = oDesign.ValidateDesign()
    print(f"Validation: {result}")
except Exception as e:
    print(f"Validation error: {e}")

# Check messages
print("\n--- Desktop Messages ---")
try:
    msgs = oDesktop.GetMessages("", "", 2)  # type 2 = all messages
    for m in msgs:
        print(f"  {m}")
except Exception as e:
    print(f"Messages error: {e}")
    try:
        msgs = oDesktop.GetMessages("Printed_Dipole_v5", "Dipole_Balun", 2)
        for m in msgs:
            print(f"  {m}")
    except Exception as e2:
        print(f"Messages error 2: {e2}")

# List objects
print("\n--- Model Objects ---")
try:
    oEditor = oDesign.SetActiveEditor("3D Modeler")
    objs = oEditor.GetObjectsInGroup("Solids")
    print(f"Solids: {objs}")
    sheets = oEditor.GetObjectsInGroup("Sheets")
    print(f"Sheets: {sheets}")
except Exception as e:
    print(f"Objects error: {e}")

# List setups
print("\n--- Analysis Setups ---")
try:
    oAnalysis = oDesign.GetModule("AnalysisSetup")
    setups = oAnalysis.GetSetups()
    print(f"Setups: {setups}")
except Exception as e:
    print(f"Setups error: {e}")

# List boundaries
print("\n--- Boundaries ---")
try:
    oBnd = oDesign.GetModule("BoundarySetup")
    bnds = oBnd.GetBoundaries()
    print(f"Boundaries: {bnds}")
    excitations = oBnd.GetExcitations()
    print(f"Excitations: {excitations}")
except Exception as e:
    print(f"Boundaries error: {e}")

# Try analyze with PyAEDT
print("\n--- Try hfss.analyze_setup ---")
try:
    hfss.analyze_setup("Setup1")
    print("analyze_setup completed!")
except Exception as e:
    print(f"analyze_setup error: {e}")
    traceback.print_exc()

# Try direct COM Analyze
print("\n--- Try oDesign.Analyze ---")
try:
    oDesign.Analyze("Setup1")
    print("Analyze completed!")
except Exception as e:
    print(f"Analyze error: {e}")
    traceback.print_exc()

print("\nTest done.")
try: hfss.release_desktop()
except: pass
