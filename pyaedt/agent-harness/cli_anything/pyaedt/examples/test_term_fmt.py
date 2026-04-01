"""Test: Use 'Terminals' section (not 'Modes') in DrivenTerminal mode!
This may be the ROOT CAUSE - we've been using DrivenModal syntax in DrivenTerminal."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "terminal_format_test.log")

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

# Test each design one at a time
results = {}

for test_name, port_args in [
    ("Terminals_format", 
     ["NAME:Port1",
      "Objects:=", ["PRect"],
      "RenormalizeAllTerminals:=", True,
      "DoDeembed:=", False,
      ["NAME:Terminals",
       ["NAME:Port1_T1",
        "TerminalResistance:=", "50ohm",
        "UseIntLine:=", True,
        ["NAME:IntLine",
         "Start:=", ["0mm", "-3mm", "1.635mm"],
         "End:=",   ["0mm", "-3mm", "-0.035mm"]]]]]),
    
    ("Terminals_noIntLine",
     ["NAME:Port1",
      "Objects:=", ["PRect"],
      "RenormalizeAllTerminals:=", True,
      "DoDeembed:=", False,
      ["NAME:Terminals",
       ["NAME:Port1_T1",
        "TerminalResistance:=", "50ohm"]]]),
    
    ("NoModesNoTerminals",
     ["NAME:Port1",
      "Objects:=", ["PRect"],
      "RenormalizeAllTerminals:=", True,
      "DoDeembed:=", False]),

    ("TermWithRefCond",
     ["NAME:Port1",
      "Objects:=", ["PRect"],
      "RenormalizeAllTerminals:=", True,
      "DoDeembed:=", False,
      ["NAME:Terminals",
       ["NAME:Port1_T1",
        "TerminalResistance:=", "50ohm",
        "UseIntLine:=", True,
        ["NAME:IntLine",
         "Start:=", ["0mm", "-3mm", "1.635mm"],
         "End:=",   ["0mm", "-3mm", "-0.035mm"]],
        "ReferenceObject:=", "Gnd"]]]),

    ("Modes_format",
     ["NAME:Port1",
      "Objects:=", ["PRect"],
      "RenormalizeAllTerminals:=", True,
      "DoDeembed:=", False,
      ["NAME:Modes",
       ["NAME:Mode1",
        "ModeNum:=", 1,
        "UseIntLine:=", True,
        ["NAME:IntLine",
         "Start:=", ["0mm", "-3mm", "1.635mm"],
         "End:=",   ["0mm", "-3mm", "-0.035mm"]],
        "CharImp:=", "Zpi"]]]),
]:
    design_name = test_name
    print(f"\n===== TEST: {test_name} =====")
    
    try:
        hfss = Hfss(
            projectname=os.path.join(PROJECT_DIR, "TermFmtTest"),
            designname=design_name,
            solution_type="DrivenTerminal",
            non_graphical=True,
            new_desktop_session=False,
            specified_version="2019.1",
        )
        hfss.modeler.model_units = "mm"
        oDesign = hfss.odesign
        oEditor = oDesign.SetActiveEditor("3D Modeler")
        oDesktop = hfss.odesktop
        oBnd = oDesign.GetModule("BoundarySetup")
        
        # Create geometry - conductors extend beyond port at y=-3
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
        
        # Port sheet at y=-3 (cuts through both conductors)
        pz_start = -0.035
        pz_end = 1.635
        port_h = pz_end - pz_start
        oEditor.CreateRectangle(
            ["NAME:RectangleParameters",
             "IsCovered:=", True,
             "XStart:=", "-1.5mm", "YStart:=", "-3mm", "ZStart:=", f"{pz_start}mm",
             "Width:=", "3mm", "Height:=", f"{port_h}mm",
             "WhichAxis:=", "Y"],
            ["NAME:Attributes", "Name:=", "PRect", "Flags:=", "", "Color:=", "(0 0 255)",
             "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
             "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
             "SolveInside:=", True, "IsMaterialEditable:=", True,
             "UseMaterialAppearance:=", False, "IsLightweight:=", False])
        
        # AirBox
        oEditor.CreateBox(
            ["NAME:BoxParameters",
             "XPosition:=", "-15mm", "YPosition:=", "-15mm", "ZPosition:=", "-10mm",
             "XSize:=", "30mm", "YSize:=", "30mm", "ZSize:=", "25mm"],
            ["NAME:Attributes", "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
             "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
             "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
             "SolveInside:=", True, "IsMaterialEditable:=", True,
             "UseMaterialAppearance:=", False, "IsLightweight:=", False])
        
        # Radiation boundary
        oBnd.AssignRadiation(
            ["NAME:Rad1",
             "Objects:=", ["Air"],
             "IsFssReference:=", False,
             "IsForPML:=", False])
        
        print(f"  Geometry OK")
        
        # Assign LumpedPort
        try:
            oBnd.AssignLumpedPort(port_args)
            print(f"  Port assigned OK")
        except Exception as e:
            print(f"  Port FAILED: {e}")
            results[test_name] = f"Port creation failed: {e}"
            continue
        
        # Check excitations
        try:
            exc = oBnd.GetExcitations()
            print(f"  Excitations: {exc}")
        except:
            print(f"  GetExcitations failed")
        
        # Setup
        oAnalysis = oDesign.GetModule("AnalysisSetup")
        oAnalysis.InsertSetup("HfssDriven",
            ["NAME:Setup1",
             "Frequency:=", "2.217GHz",
             "MaxDeltaS:=", 0.05,
             "MaximumPasses:=", 3,
             "MinimumPasses:=", 1,
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
        
        # Validate
        v = oDesign.ValidateDesign()
        print(f"  Validate: {v}")
        
        # Get messages
        msgs = oDesktop.GetMessages("", "", 2)
        term_msgs = [m for m in msgs if "terminal" in str(m).lower() or "port" in str(m).lower()]
        for m in term_msgs:
            print(f"  MSG: {m}")
        
        if v:
            print(f"  -> VALIDATION PASSED! Trying analyze...")
            try:
                oDesign.Analyze("Setup1")
                print(f"  -> SIMULATION SUCCEEDED!")
                results[test_name] = "SUCCESS"
            except Exception as e:
                print(f"  -> Analyze failed: {e}")
                results[test_name] = f"Analyze failed: {e}"
        else:
            results[test_name] = "Validation failed (no terminals?)"
            
    except Exception as e:
        print(f"  EXCEPTION: {e}")
        traceback.print_exc()
        results[test_name] = f"Exception: {e}"

print(f"\n\n===== RESULTS SUMMARY =====")
for k, v in results.items():
    print(f"  {k}: {v}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
