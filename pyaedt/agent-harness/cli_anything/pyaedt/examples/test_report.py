"""Test report creation in AEDT 2019.1 after successful simulation.
The simulation WORKS! Just need correct CreateReport syntax."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "report_test.log")

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

# Reopen the existing project that already has simulation results
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "DefWP"),
    designname="T1",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
oDesign = hfss.odesign
oDesktop = hfss.odesktop
oBnd = oDesign.GetModule("BoundarySetup")
oReport = oDesign.GetModule("ReportSetup")

print("=== Trying different CreateReport formats ===")

# Get terminal name
terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
terminal = str(terms[0]) if terms else "Trace_T1"

# Format 1: 7-arg (used in newer AEDT)
print("\n--- Format 1: 7 args ---")
try:
    oReport.CreateReport(
        "S11_1", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal},{terminal}))"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Format 2: 8-arg (some AEDT versions need context)
print("\n--- Format 2: 8 args with context ---")
try:
    oReport.CreateReport(
        "S11_2", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        [],  # context
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal},{terminal}))"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Format 3: With NAME syntax
print("\n--- Format 3: NAME syntax ---")
try:
    oReport.CreateReport(
        "S11_3", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["NAME:Context"],
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal},{terminal}))"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Format 4: Minimal
print("\n--- Format 4: Create, then add trace ---")
try:
    oReport.CreateReport(
        "S11_4", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["NAME:Context"],
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"St({terminal},{terminal})"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Format 5: Older CreateReport with different arg count
print("\n--- Format 5: 6 args ---")
try:
    oReport.CreateReport(
        "S11_5", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal},{terminal}))"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Format 6: Try Modal Solution Data instead of Terminal
print("\n--- Format 6: Modal Solution Data ---")
try:
    oReport.CreateReport(
        "S11_6", "Modal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        [],
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", ["dB(S(Port1,Port1))"]])
    print("  OK!")
except Exception as e:
    print(f"  Failed: {e}")

# Try GetAllReportNames
print("\n--- Available reports ---")
try:
    reports = oReport.GetAllReportNames()
    print(f"  Reports: {reports}")
except Exception as e:
    print(f"  GetAllReportNames: {e}")

# Try to export any successful report
for name in ["S11_1", "S11_2", "S11_3", "S11_4", "S11_5", "S11_6"]:
    try:
        csv_path = os.path.join(PROJECT_DIR, f"{name}.csv")
        oReport.ExportToFile(name, csv_path)
        print(f"  Exported {name} to {csv_path}")
        # Read and print
        with open(csv_path, 'r') as f:
            content = f.read()
        print(f"  Content ({len(content)} chars):")
        print(f"  {content[:500]}")
        break
    except Exception as e:
        print(f"  Export {name}: {e}")

# Try GetSolutionDataPerVariation
print("\n--- Direct solution data ---")
try:
    data = oReport.GetSolutionDataPerVariation(
        "Terminal Solution Data",
        "Setup1 : LastAdaptive",
        ["NAME:Context"],
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        [f"St({terminal},{terminal})"])
    print(f"  Data: {data}")
except Exception as e:
    print(f"  Failed: {e}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
