"""Test: Use Split/Section to create topologically-connected port surfaces.
The idea: split the Trace at the port location so the internal face
becomes part of the conductor topology, enabling terminal detection."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "split_test.log")

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

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "SplitTest"),
    designname="T1",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oDesktop = hfss.odesktop

# Geometry
print("Creating geometry...")
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", "-5mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-15mm", "YPosition:=", "-15mm", "ZPosition:=", "-10mm",
     "XSize:=", "30mm", "YSize:=", "30mm", "ZSize:=", "25mm"],
    ["NAME:Attributes", "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  Geometry OK")

oBnd = oDesign.GetModule("BoundarySetup")

# ---- Approach A: Section the Trace at port location ----
print("\n--- Approach A: Section the substrate to create port face ---")
# Section creates a sheet object from the cross-section of a 3D object
# We section at the y=-5mm plane (substrate edge)
# Actually, we need a port sheet that spans from Gnd to Trace
# Let's create it by sectioning the substrate

try:
    # Section the substrate at y = -5mm
    oEditor.Section(
        ["NAME:Selections", "Selections:=", "Sub", "NewPartsModelFlag:=", "Model"],
        ["NAME:SectionToParameters",
         "CreateNewObjects:=", True,
         "SectionPlane:=", "YZ",  # Y is the split direction? or use plane?
         "SplitCrossingObjectsOnly:=", False,
         "DeleteInvalidObjects:=", True])
    print("  Section Sub succeeded!")
    
    # Check objects
    objs = oEditor.GetObjectsInGroup("Sheets")
    print(f"  Sheets after section: {objs}")
except Exception as e:
    print(f"  Section failed: {e}")

# ---- Approach B: Split Trace to create internal face ----
print("\n--- Approach B: Split Trace at y=-5 plane ---")
try:
    oEditor.Split(
        ["NAME:Selections", "Selections:=", "Trace", "NewPartsModelFlag:=", "Model"],
        ["NAME:SplitToParameters",
         "SplitPlane:=", "XZ",
         "SplitCrossingObjectsOnly:=", False,
         "DeleteInvalidObjects:=", True])
    print("  Split Trace at Y=0 succeeded!")
    
    objs = oEditor.GetObjectsInGroup("Solids")
    print(f"  Solids after split: {objs}")
    
    # After splitting, Trace should be split into two halves
    # The split face at y=0 should be shared by both halves
except Exception as e:
    print(f"  Split Trace failed: {e}")

# ---- Approach C: Create port sheet INSIDE the substrate ----
# Make it slightly overlap with conductor faces
print("\n--- Approach C: Port sheet overlapping conductors ---")
try:
    # Create a rect that EXTENDS INTO the conductor volumes
    # This forces topological connection
    oEditor.CreateRectangle(
        ["NAME:RectangleParameters",
         "IsCovered:=", True,
         "XStart:=", "-1.5mm", "YStart:=", "-5mm", "ZStart:=", "-0.04mm",
         "Width:=", "3mm", "Height:=", "1.68mm",
         "WhichAxis:=", "Y"],
        ["NAME:Attributes", "Name:=", "PRect", "Flags:=", "", "Color:=", "(0 0 255)",
         "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
         "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
         "SolveInside:=", True, "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False, "IsLightweight:=", False])
    print("  PRect created (slightly extending into conductors)")
    
    # Now subtract the port rect from Gnd and Trace to create connectivity
    print("  Subtracting PRect from conductors to create topology...")
    try:
        oEditor.Subtract(
            ["NAME:Selections",
             "Blank Parts:=", "Gnd", "Tool Parts:=", "PRect"],
            ["NAME:SubtractParameters",
             "KeepOriginals:=", True])
        print("  Subtract from Gnd succeeded!")
    except Exception as e:
        print(f"  Subtract from Gnd failed: {e}")
    
    try:
        oEditor.Subtract(
            ["NAME:Selections",
             "Blank Parts:=", "Trace", "Tool Parts:=", "PRect"],
            ["NAME:SubtractParameters",
             "KeepOriginals:=", True])
        print("  Subtract from Trace succeeded!")
    except Exception as e:
        print(f"  Subtract from Trace failed: {e}")
    
    # Now assign lumped port to PRect
    oBnd.AssignLumpedPort(
        ["NAME:Port1",
         "Objects:=", ["PRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0mm", "-5mm", "1.635mm"],
            "End:=",   ["0mm", "-5mm", "-0.035mm"]],
           "CharImp:=", "Zpi"]]])
    print("  Port1 assigned on PRect")
    
except Exception as e:
    print(f"  Approach C failed: {e}")
    traceback.print_exc()

# ---- Approach D: Use AutoIdentifyPorts with face of Trace ----
print("\n--- Approach D: AutoIdentifyPorts on Trace face ---")
try:
    trace_faces = oEditor.GetFaceIDs("Trace")
    print(f"  Trace faces: {trace_faces}")
    
    for fid_str in trace_faces:
        fid = int(fid_str)
        try:
            oBnd.AutoIdentifyPorts(
                ["NAME:Faces", fid],
                True,
                ["NAME:ReferenceConductors", "Gnd"],
                f"AP_{fid}",
                True)
            print(f"  AutoIdentifyPorts face {fid} SUCCESS!")
        except Exception as e:
            pass  # silently skip
    
    exc = oBnd.GetExcitations()
    print(f"  Excitations after AutoID: {exc}")
except Exception as e:
    print(f"  Approach D failed: {e}")

# ---- Approach E: Use AutoIdentifyPorts on Gnd face ----
print("\n--- Approach E: AutoIdentifyPorts on Gnd faces ---")
try:
    gnd_faces = oEditor.GetFaceIDs("Gnd")
    print(f"  Gnd faces: {gnd_faces}")
    
    for fid_str in gnd_faces:
        fid = int(fid_str)
        try:
            oBnd.AutoIdentifyPorts(
                ["NAME:Faces", fid],
                True,
                ["NAME:ReferenceConductors", "Gnd"],
                f"GP_{fid}",
                True)
            print(f"  AutoIdentifyPorts Gnd face {fid} SUCCESS!")
        except Exception as e:
            pass  # silently skip
    
    exc = oBnd.GetExcitations()
    print(f"  Excitations after Gnd AutoID: {exc}")
except Exception as e:
    print(f"  Approach E failed: {e}")

# Setup and validate
print("\n--- Setup & Validate ---")
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])

oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", "2.217GHz",
     "MaxDeltaS:=", 0.05,
     "MaximumPasses:=", 6,
     "MinimumPasses:=", 2,
     "MinimumConvergedPasses:=", 1,
     "PercentRefinement:=", 30,
     "IsEnabled:=", True,
     "BasisOrder:=", 1,
     "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True,
     "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False,
     "Target:=", 0.3333])

hfss.save_project()
print("  Saved")

v = oDesign.ValidateDesign()
print(f"  Validation: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    print(f"  MSG: {m}")

# Try analyze if validation passed
if v != 0:
    print("\n--- Analyze ---")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  FAILED ({time.time()-t0:.1f}s): {e}")
else:
    print("\n  Validation failed, trying Analyze anyway...")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  FAILED ({time.time()-t0:.1f}s): {e}")
    
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if "error" in str(m).lower():
            print(f"  {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
