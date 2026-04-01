"""Minimal test: can AEDT 2019.1 still run a simulation?"""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

# patches
try:
    from pyaedt import desktop as _dm
    _o = _dm.Desktop.__init__
    def _p(self, *a, **kw):
        _o(self, *a, **kw)
        if not hasattr(self, 'student_version'): self.student_version = False
    _dm.Desktop.__init__ = _p
except: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(self, app):
        try: _o2(self, app)
        except AttributeError:
            self._app = app; self.design_settings = None; self.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except: pass

from pyaedt import Hfss

PROJECT_DIR = r"D:\class_design"
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "TestMin2"),
    designname="Test1",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=True,
    specified_version="2019.1",
)
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oBnd = oDesign.GetModule("BoundarySetup")
hfss.modeler.model_units = "mm"

# Simple dipole in air
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-35mm", "YPosition:=", "-1.5mm", "ZPosition:=", "0mm",
     "XSize:=", "34mm", "YSize:=", "3mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes",
     "Name:=", "Arm1", "Flags:=", "", "Color:=", "(255 0 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "1mm", "YPosition:=", "-1.5mm", "ZPosition:=", "0mm",
     "XSize:=", "34mm", "YSize:=", "3mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes",
     "Name:=", "Arm2", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# AirBox
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-70mm", "YPosition:=", "-40mm", "ZPosition:=", "-35mm",
     "XSize:=", "140mm", "YSize:=", "80mm", "ZSize:=", "70mm"],
    ["NAME:Attributes",
     "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oBnd.AssignRadiation(
    ["NAME:Rad1", "Objects:=", ["Air"], "IsFssReference:=", False, "IsForPML:=", False])
print("Radiation BC OK")

# AutoIdentifyPorts on y_min face
air_faces = oEditor.GetFaceIDs("Air")
print(f"AirBox faces: {air_faces}")

# Find y_min face by checking face centers
for i, fid in enumerate(air_faces):
    try:
        center = oEditor.GetFaceCenter(int(fid))
        print(f"  Face {i} (ID={fid}): center=({center[0]}, {center[1]}, {center[2]})")
    except:
        print(f"  Face {i} (ID={fid}): cannot get center")

# Use y_min face (should have most negative y center)
# For now try face index 4 (common for y_min in AEDT)
# Actually let me find it programmatically
import operator
face_centers = []
for fid in air_faces:
    try:
        c = oEditor.GetFaceCenter(int(fid))
        face_centers.append((int(fid), float(c[0]), float(c[1]), float(c[2])))
    except:
        pass

# y_min face has the most negative Y center
y_min_face = min(face_centers, key=lambda x: x[2])
print(f"y_min face: ID={y_min_face[0]}, center_y={y_min_face[2]}")

oBnd.AutoIdentifyPorts(
    ["NAME:Faces", y_min_face[0]], True,
    ["NAME:ReferenceConductors", "Arm1"],
    "Port1", True)
print("AutoIdentifyPorts OK")

terms = oBnd.GetExcitationsOfType("Terminal")
print(f"Terminals: {terms}")

# Setup
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", "2.217GHz",
     "MaxDeltaS:=", 0.05, "MaximumPasses:=", 6,
     "MinimumPasses:=", 2, "MinimumConvergedPasses:=", 1,
     "PercentRefinement:=", 30, "IsEnabled:=", True,
     "BasisOrder:=", 1, "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True, "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False, "Target:=", 0.3333])
print("Setup OK")

# Validate
hfss.save_project()
v = oDesign.ValidateDesign()
print(f"Validation: {v}")

if v:
    print("Analyzing...")
    oDesign.Analyze("Setup1")
    print("SIMULATION SUCCESS!")
else:
    print("Validation FAILED, trying anyway...")
    try:
        oDesign.Analyze("Setup1")
        print("SIMULATION SUCCESS! (despite val fail)")
    except Exception as e:
        print(f"ANALYZE FAILED: {e}")

hfss.save_project()
print("DONE")
