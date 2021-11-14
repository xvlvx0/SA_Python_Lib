# -*- coding: utf-8 -*-
"""
SAPyLib = Spatial Analyzer Python .NET library

This library provides the communication layer to the SA SDK .NET dll.
Author: L. Ververgaard
"""
import sys
import os
import logging
import clr

clr.AddReference("System")
clr.AddReference("System.Reflection")
import System
import System.Reflection
from System import Array, Double


# SA SDK dll
basepath = "C:/Analyzer Data/Scripts/SAPython"
sa_sdk_dll_file = os.path.join(basepath, "dll/Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
NrkSdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = NrkSdk.Connect("127.0.0.1")
if not SAConnected:
    raise IOError("Connection to SA SDK failed!")


# Get the logger
log = logging.getLogger(__name__)


# ###############
# Base methods ##
# ###############
def getResult_Bare(func_name):
    """Get the methods execution result without processing."""
    boolean, result = NrkSdk.GetMPStepResult(0)
    log.debug(f"{func_name}: {boolean}, {result}")
    return boolean, result


def getResult(func_name):
    """Get the methods execution result and process the result."""
    boolean, result = NrkSdk.GetMPStepResult(0)
    if result == -1:
        # SDKERROR = -1
        log.error(f"{func_name}: {boolean}, {result}")
        MPStepMessages()
        raise SystemError("Execution raised: SDKERROR!")
    elif result == 0:
        # UNDONE = 0
        log.warning(f"{func_name}: {boolean}, {result}")
        log.warning("Execution: undone.")
        MPStepMessages()
        raise SystemError("Execution raised: UNDONE!")
    elif result == 1:
        # INPROGRESS = 1
        log.info(f"{func_name}: {boolean}, {result}")
        MPStepMessages()
        log.warning("Execution: inprogress.")
        return True
    elif result == 2:
        # DONESUCCESS = 2
        return True
    elif result == 3:
        # DONEFATALERROR = 3
        log.error(f"{func_name}: {boolean}, {result}")
        log.error("Execution: FAILED!")
        MPStepMessages()
        return False
    elif result == 4:
        # DONEMINORERROR = 4
        log.warning(f"{func_name}: {boolean}, {result}")
        log.warning("Execution: FAILED - minor error!")
        MPStepMessages()
        return False
    elif result == 5:
        # CURRENTTASK = 5
        log.debug(f"{func_name}: {boolean}, {result}")
        log.info("Execution: current task")
        MPStepMessages()
        return True
    elif result == 6:
        # UNKNOWN = 6
        log.debug(f"{func_name}: {boolean}, {result}")
        log.info("I have no clue!")
        MPStepMessages()
        raise SystemError("I have no clue!")


def MPStepMessages():
    """Get the MPStep messages."""
    log.debug("Get the MPStep messages.")
    stringList = System.Runtime.InteropServices.VariantWrapper([])
    try:
        vStringList = NrkSdk.GetMPStepMessages(stringList)
        if vStringList[0]:
            messages = []
            for i in range(vStringList[1].GetLength(0)):
                messages.append(vStringList[1][i])
                log.info(f"MPStepMessage: {vStringList[1][i]}")
            return messages
        else:
            return False
    except System.Runtime.InteropServices.COMException as err:
        log.error(f"Getting MP Step failed with error: {err}")
        return False


class Point3D:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z


class NamedPoint:
    def __init__(self, point):
        self.collection = point[0]
        self.group = point[1]
        self.name = point[2]


class NamedPoint3D:
    def __init__(self, point):
        log.debug("NamedPoint3D")

        if not isinstance(point, NamedPoint):
            self.collection = point[0]
            self.group = point[1]
            self.name = point[2]
        else:
            self.collection = point.collection
            self.group = point.group
            self.name = point.name

        # Get XYZ
        pData = get_point_coordinate(self.collection, self.group, self.name)
        log.info(f"pData: {pData}")
        self.X = pData[0]
        self.Y = pData[1]
        self.Z = pData[2]


def FahrenheitToCelsius(tempF):
    """Convert Fahrenheit to Celsius."""
    tempC = (tempF - 32) * (5 / 9)
    return tempC


def InchHgtoMilliBar(pressInchHg):
    """Convert Inch Mercury (Hg) to mBar."""
    pressMilliBar = pressInchHg * 33.864
    return pressMilliBar


def PythonToCSharp2Darray(input_list, array_depth=(4, 4)):
    """Convert a python N-List to a C# Array."""
    a = Array.CreateInstance(Double, array_depth[0], array_depth[1])
    for i in range(array_depth[0]):
        for j in range(array_depth[1]):
            a.SetValue(input_list[i][j], i, j)
    return a


def CSharpToPython2Darray(input_array, array_depth=(4, 4)):
    """Convert a C# Array to a N-List."""
    a = []
    a.append([])
    for i in range(array_depth[0]):
        for j in range(array_depth[1]):
            a[i].append(input_array.GetValue(i, j))
        a.append([])
    return a


# ##############################
# Chapter 2 - File Operations ##
# ##############################
def find_files_in_directory(directory, searchPattern):
    """p27"""
    func_name = "Find Files in Directory"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Directory", directory)
    NrkSdk.SetStringArg("File Name Pattern", searchPattern)
    NrkSdk.SetBoolArg("Recursive?", False)
    NrkSdk.ExecuteStep()
    results = getResult(func_name)
    if not results:
        return False
    stringList = System.Runtime.InteropServices.VariantWrapper([])
    vStringList = NrkSdk.GetStringRefListArg("Files", stringList)
    if vStringList[0]:
        files = []
        for i in range(vStringList[1].GetLength(0)):
            files.append(vStringList[1][i])
        return files
    else:
        return False


# ######################################
# Chapter 3 - Process Flow Operations ##
# ######################################
def object_existence_test(ColObjName):
    """p111"""
    func_name = ""
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    NrkSdk.ExecuteStep()
    getResult(func_name)


def ask_for_string(question, initialanswer=""):
    """p117"""
    func_name = "Ask for String"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Question to ask", question)
    NrkSdk.SetBoolArg("Password Entry?", False)
    NrkSdk.SetStringArg("Initial Answer", initialanswer)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    answer = NrkSdk.GetStringArg("Answer", "")
    return answer[1]


def ask_for_string_pulldown(question, answers):
    """p118"""
    func_name = "Ask for String (Pull-Down Version)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    # question section
    questionList = [
        question,
    ]
    QvStringList = System.Runtime.InteropServices.VariantWrapper(questionList)
    NrkSdk.SetStringRefListArg("Question or Statement", QvStringList)
    # answers section
    AvStringList = System.Runtime.InteropServices.VariantWrapper(answers)
    NrkSdk.SetStringRefListArg("Possible Answers", AvStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    answer = NrkSdk.GetStringArg("Answer", "")
    log.debug(f"Answer: {answer}")
    return answer[1]


def ask_for_user_decision_extended(question, choices):
    """p122"""
    func_name = "Ask for User Decision Extended"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    stringList = [
        question,
    ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Question or Statement", vStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    # NrkSdk.NOT_SUPPORTED("Button Answers")
    # NrkSdk.NOT_SUPPORTED("Step to jump to if Canceled (-1 will fail step on Cancel)")
    raise SystemError("Not Implemented by SDK!")


# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################


# ###########################
# Chapter 5 - View Control ##
# ###########################
def hide_objects(collection, name, objtype):
    """p153"""
    func_name = "Hide Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = [
        f"{collection}::{name}::{objtype}",
    ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Hide", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def show_hide_objects(collection, objtype, hide):
    """p154"""
    # Hide = hide/True
    # Show = hide/False
    func_name = "Show / Hide by Object Type"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetBoolArg("All Collections?", False)
    NrkSdk.SetCollectionNameArg("Specific Collection", collection)
    # Available options:
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse",
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter",
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh",
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group",
    NrkSdk.SetObjectTypeArg("Object Type To Show / Hide", objtype)
    log.debug(f"{collection}::{objtype} Hide? {hide}")
    NrkSdk.SetBoolArg("Hide? (Show = FALSE)", hide)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def show_objects(collection, objects, name):
    """p155"""
    func_name = "Show Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = [
        f"{collection}::{objects}::{name}",
    ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Show", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def show_hide_callout_view(collection, calloutname, show):
    """p159"""
    # Show = show/True
    # Hide = show/False
    func_name = "Show / Hide Callout View"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Callout View To Show", collection, calloutname)
    NrkSdk.SetBoolArg("Show Callout View?", show)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def hide_all_callout_views():
    """p160"""
    func_name = "Hide All Callout Views"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def center_graphics_about_objects(ColWild, ObjWild):
    """p170"""
    func_name = ""
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ######################################
# Chapter 6 - Cloud Viewer Operations ##
# ######################################


# ######################################
# Chapter 7 - Construction Operations ##
# ######################################
def rename_point(orgCol, orgGrp, orgName, newCol, newGrp, newName, **kwargs):
    """p192"""
    func_name = "Rename Point"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Original Point Name", orgCol, orgGrp, orgName)
    NrkSdk.SetPointNameArg("New Point Name", newCol, newGrp, newName)
    NrkSdk.SetBoolArg("Overwrite if exists?", kwargs.get("overwrite", False))
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        print("\n\nERROR: renaming failed.\n")
        sys.exit(1)


def rename_collection(fromName, toName):
    """p194"""
    func_name = "Rename Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Original Collection Name", fromName)
    NrkSdk.SetCollectionNameArg("New Collection Name", toName)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        raise SystemError("Renaming the folder failed!")


def rename_object(old_col, old_name, new_col, new_name):
    """p195"""
    func_name = "Rename Object"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Original Object Name", old_col, old_name)
    NrkSdk.SetCollectionObjectNameArg("New Object Name", new_col, new_name)
    NrkSdk.SetBoolArg("Overwrite if exists?", False)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        raise SystemError("Renaming the object failed!")


def delete_points(collection, group, point):
    """p197"""
    func_name = "Delete Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    ptNameList = [
        f"{collection}::{group}::{point}",
    ]
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg("Point Names", vPointObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_points_wildcard_selection(Col, Group, ObjType, pName):
    """p198"""
    func_name = "Delete Points WildCard Selection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = [
        f"{Col}::{Group}::{ObjType}",
    ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Groups to Delete From", vObjectList)
    NrkSdk.SetPointNameArg("WildCard Selection Names", "*", "*", pName)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_objects_from_surface_faces_runtime_select(*args, **kwargs):
    """p199"""
    func_name = "Construct Objects From Surface Faces - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    if "facetype" in kwargs:
        facetype = kwargs["facetype"]
        if not facetype:
            log.warning("No facetype selected!")
            sys.exit(1)
    # Set correct states
    state = [False, False, False, False, False, False, False]
    if facetype == "plane":
        state[0] = True
    elif facetype == "cylinder":
        state[1] = True
    elif facetype == "spere":
        state[2] = True
    elif facetype == "cone":
        state[3] = True
    elif facetype == "line":
        state[4] = True
    elif facetype == "point":
        state[5] = True
    elif facetype == "circle":
        state[6] = True
    # Set the correct type
    NrkSdk.SetBoolArg("Construct Planes?", state[0])
    NrkSdk.SetBoolArg("Construct Cylinders?", state[1])
    NrkSdk.SetBoolArg("Construct Spheres?", state[2])
    NrkSdk.SetBoolArg("Construct Cones?", state[3])
    NrkSdk.SetBoolArg("Construct Lines?", state[4])
    NrkSdk.SetBoolArg("Construct Points?", state[5])
    NrkSdk.SetBoolArg("Construct Circles?", state[6])
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_or_construct_default_collection(colName):
    """p201"""
    func_name = "Set (or construct) default collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection Name", colName)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_collection(name, makeDefault):
    """p202"""
    # default collection = makeDefault/True
    # collection = makeDefault/False
    func_name = "Construct Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection Name", name)
    NrkSdk.SetStringArg("Folder Path", "")
    NrkSdk.SetBoolArg("Make Default Collection?", makeDefault)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_active_collection():
    """p203"""
    func_name = "Get Active Collection Name"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetCollectionNameArg("Currently Active Collection Name", "")
    if sValue[0]:
        return sValue[1]
    else:
        return False


def delete_collection(collection):
    """p204"""
    func_name = "Delete Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_a_point(collection, group, name, x, y, z):
    """p208"""
    func_name = "Construct a Point in Working Coordinates"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.SetVectorArg("Working Coordinates", x, y, z)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_point_at_intersection_of_plane_and_line(
    collectionplane, plane, collectionline, line, collectionpoint, grouppoint, point
):
    """p218"""
    func_name = "Construct Point at Intersection of Plane and Line"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", collectionplane, plane)
    NrkSdk.SetCollectionObjectNameArg("Line Name", collectionline, line)
    NrkSdk.SetPointNameArg("Resulting Point Name", collectionpoint, grouppoint, point)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_line_2_points(name, fCol, fGrp, fTarg, sCol, sGrp, sTarg):
    """p262"""
    func_name = "Construct Line 2 Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Line Name", "", name)
    NrkSdk.SetPointNameArg("First Point", fCol, fGrp, fTarg)
    NrkSdk.SetPointNameArg("Second Point", sCol, sGrp, sTarg)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_plane(planeCol, planeName):
    """p271"""
    func_name = "Construct Plane"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", planeCol, planeName)
    NrkSdk.SetVectorArg("Plane Center (in working coordinates)", 0.0, 0.0, 0.0)
    NrkSdk.SetVectorArg("Plane Normal (in working coordinates)", 0.0, 0.0, 1.0)
    NrkSdk.SetDoubleArg("Plane Edge Dimension", 0.0)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_frame_known_origin_object_direction_object_direction(
    ptCol,
    ptGrp,
    ptTarg,
    x,
    y,
    z,
    priCol,
    priObj,
    priAxisDef,
    secCol,
    secObj,
    secAxisDef,
    frameCol,
    frameName,
):
    """p327"""
    func_name = "Construct Frame, Known Origin, Object Direction, Object Direction"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Known Point", ptCol, ptGrp, ptTarg)
    NrkSdk.SetVectorArg("Known Point Value in New Frame", x, y, z)
    NrkSdk.SetCollectionObjectNameArg("Primary Axis Object", priCol, priObj)
    # Available options:
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis",
    # "-Z Axis",
    NrkSdk.SetAxisNameArg("Primary Axis Defines Which Axis", priAxisDef)
    NrkSdk.SetCollectionObjectNameArg("Secondary Axis Object", secCol, secObj)
    # Available options:
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis",
    # "-Z Axis",
    NrkSdk.SetAxisNameArg("Secondary Axis Defines Which Axis", secAxisDef)
    NrkSdk.SetCollectionObjectNameArg("Frame Name (Optional)", frameCol, frameName)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def create_relationship_callout(collection, calloutname, relationshipcollection, relationshipname, *args, **kwargs):
    """p363"""
    func_name = "Create Relationship Callout"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection, calloutname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relationshipcollection, relationshipname)
    # get X position
    if "xpos" in kwargs:
        xpos = kwargs["xpos"]
    else:
        xpos = 0.1
    # get Y position
    if "ypos" in kwargs:
        ypos = kwargs["ypos"]
    else:
        ypos = 0.1
    # set values
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    vStringList = System.Runtime.InteropServices.VariantWrapper([])
    NrkSdk.SetEditTextArg("Additional Notes (blank for none)", vStringList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def create_text_callout(collection, calloutname, text, *args, **kwargs):
    """p365"""
    func_name = "Create Text Callout"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection, calloutname)
    stringList = [
        text,
    ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Text", vStringList)
    # get X position
    if "xpos" in kwargs:
        xpos = kwargs["xpos"]
    else:
        xpos = 0.1
    # get Y position
    if "ypos" in kwargs:
        ypos = kwargs["ypos"]
    else:
        ypos = 0.1
    # set values
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    NrkSdk.SetPointNameArg("Callout Anchor Point (Optional)", "", "", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_default_callout_view_properties(calloutname):
    """p372"""
    func_name = "Set Default Callout View Properties"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Default Callout View Name", calloutname)
    NrkSdk.SetBoolArg("Lock View Point?", True)
    NrkSdk.SetBoolArg("Recall Working Frame?", True)
    NrkSdk.SetBoolArg("Recall Visible Layer?", True)
    NrkSdk.SetIntegerArg("Callout Leader Thickness", 2)
    NrkSdk.SetColorArg("Callout Leader Color", 128, 128, 128)
    NrkSdk.SetIntegerArg("Callout Border Thickness", 2)
    NrkSdk.SetColorArg("Callout Border Color", 0, 0, 255)
    NrkSdk.SetBoolArg("Divide Text with Lines?", False)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def make_a_system_string(option):
    """p390"""
    func_name = "Make a System String"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    # Available options:
    # "SA Version", "XIT Filename", "MP Filename", "MP Filename (Full Path)", "Date & Time",
    # "Date", "Date (Short)", "Time", "Key Serial Number", "Company Name",
    # "User Name",
    NrkSdk.SetSystemStringArg("String Content", option)
    NrkSdk.SetStringArg("Format String (Optional)", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetStringArg("Resultant String", "")
    if not sValue[0]:
        return False
    return sValue[1]


def make_a_point_name_runtime_select(txt):
    """p398"""
    func_name = "Make a Point Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", txt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    point = NrkSdk.GetPointNameArg("Resultant Point Name", "", "", "")
    log.debug(f"Point: {point}")
    if point[0]:
        sCol = point[1]
        sGrp = point[2]
        sTarg = point[3]
        return (sCol, sGrp, sTarg)
    else:
        return False


def make_a_point_name_ref_list_from_a_group(collection, group):
    """p401"""
    func_name = "Make a Point Name Ref List From a Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Group Name", collection, group)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if ptList[0]:
        points = []
        for i in range(ptList[1].GetLength(0)):
            points.append(NamedPoint(ptList[1][i].split("::")))
        return points
    else:
        return False


def make_a_point_name_ref_list_runtime_select(prompt):
    """p402"""
    func_name = "Make a Point Name Ref List - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", prompt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if ptList[0]:
        points = []
        for i in range(ptList[1].GetLength(0)):
            points.append(ptList[1][i].split("::"))
        return points
    else:
        return False


def make_a_collection_name_runtime_select(txt):
    """p408"""
    func_name = "Make a Collection Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", txt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetCollectionNameArg("Resultant Collection Name", "")
    if sValue[0]:
        return sValue[1]
    else:
        return False


def make_a_collection_object_name_runtime_select(prompt, obj_type):
    """p413"""
    func_name = "Make a Collection Object Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", prompt)
    # Available options:
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse",
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter",
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh",
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group",
    NrkSdk.SetObjectTypeArg("Object Type", obj_type)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    result = NrkSdk.GetCollectionObjectNameArg("Resultant Collection Object Name", "", "")
    if not result[0]:
        return False
    sCol = result[1]
    sObj = result[2]
    return (sCol, sObj)


def make_a_collection_object_name_ref_list_by_type(collection, objtype):
    """p416"""
    func_name = "Make a Collection Object Name Ref List - By Type"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Collection", collection)
    # Available options:
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse",
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter",
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh",
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group",
    NrkSdk.SetObjectTypeArg("Object Type", objtype)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userObjectList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Collection Object Name List", userObjectList)
    if objectList[0]:
        objects = []
        for i in range(objectList[1].GetLength(0)):
            objects.append(objectList[1][i].split("::"))
        return objects
    else:
        return False


def make_a_relationship_reference_list_wildCard_selection(collection, relationship):
    """p427"""
    func_name = "Make a Relationship Reference List- WildCard Selection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Collection Wildcard Criteria", collection)
    NrkSdk.SetStringArg("Relationship Wildcard Criteria", relationship)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userObjectList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Relationship Reference List", userObjectList)
    if objectList[0]:
        objects = []
        for i in range(objectList[1].GetLength(0)):
            objects.append(objectList[1][i].split("::"))
        return objects
    else:
        return False


def make_a_relationship_reference_list_runtime_selection(question) -> list:
    """p428"""
    func_name = "Make a Relationship Reference List- Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", question)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    # CStringArray objNameList;
    # SDKHelper helper(NrkSdk);
    # helper.GetCollectionObjectNameRefListArgHelper("Resultant Relationship Reference List", objNameList);
    userObjectList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Relationship Reference List", userObjectList)
    if objectList[0]:
        objects = []
        for i in range(objectList[1].GetLength(0)):
            objects.append(objectList[1][i].split("::"))
        return objects
    else:
        return False


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def get_number_of_collections():
    """p465"""
    func_name = "Get Number of Collections"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    n = NrkSdk.GetIntegerArg("Total Count", 0)
    log.debug(f"n: {n}")
    if n[0]:
        return n[1]
    else:
        return False


def get_ith_collection_name(i):
    """p466"""
    func_name = "Get i-th Collection Name"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetIntegerArg("Collection Index", i)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    collection = NrkSdk.GetCollectionNameArg("Resultant Name", "")
    log.debug(f"Collection: {collection}")
    if collection[0]:
        return collection[1]
    else:
        return False


def get_vector_group_properties(collection, vectorgroup):
    """p481"""
    func_name = "Get Vector Group Properties"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Vector Group Name", collection, vectorgroup)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    results = {}
    # For each 'Get___' request a tuple is returned, we need to save the second argument ([1]) only.
    results["total_vectors"] = NrkSdk.GetIntegerArg("Total Vectors", 0)[1]
    results["vectors_in_tolerance"] = NrkSdk.GetIntegerArg("Vectors In Tolerance", 0)[1]
    results["vectors_out_of_tolerance"] = NrkSdk.GetIntegerArg("Vectors Out Of Tolerance", 0)[1]
    results["%_vector_in_tolerance"] = NrkSdk.GetDoubleArg("% Vectors In Tolerance", 0.0)[1]
    results["%_vector_out_tolerance"] = NrkSdk.GetDoubleArg("% Vectors Out Of Tolerance", 0.0)[1]
    results["abs_max_mag"] = NrkSdk.GetDoubleArg("Absolute Max Magnitude", 0.0)[1]
    results["abs_min_mag"] = NrkSdk.GetDoubleArg("Absolute Min Magnitude", 0.0)[1]
    results["max_mag"] = NrkSdk.GetDoubleArg("Max Magnitude", 0.0)[1]
    results["min_mag"] = NrkSdk.GetDoubleArg("Min Magnitude", 0.0)[1]
    results["standard_deviation"] = NrkSdk.GetDoubleArg("Standard Deviation", 0.0)[1]
    results["standard_deviation_to_mean_zero"] = NrkSdk.GetDoubleArg("Standard Deviation Mean Zero", 0.0)[1]
    results["avg_mag"] = NrkSdk.GetDoubleArg("Avg Magnitude", 0.0)[1]
    results["avg_abs_mag"] = NrkSdk.GetDoubleArg("Avg of Abs Magnitude", 0.0)[1]
    results["high_tol"] = NrkSdk.GetDoubleArg("High Tolerance Value", 0.0)[1]
    results["low_tol"] = NrkSdk.GetDoubleArg("Low Tolerance Value", 0.0)[1]
    results["rms_val"] = NrkSdk.GetDoubleArg("RMS Value", 0.0)[1]
    return results


def set_vector_group_colorization_options_selected(collection, vectorgroup, *args, **kwargs):
    """p485"""
    func_name = "Set Vector Group Colorization Options (Selected)"
    log.debug(func_name)
    NrkSdk.SetStep("Set Vector Group Colorization Options (Selected)")
    vgNameList = [
        f"{collection}::{vectorgroup}",
    ]
    vVectorGroupNameList = System.Runtime.InteropServices.VariantWrapper(vgNameList)
    NrkSdk.SetCollectionVectorGroupNameRefListArg("Vector Groups to be Set", vVectorGroupNameList)
    NrkSdk.SetColorizationOptionsArg(
        "Colorization Options",
        kwargs["color_profile"],
        kwargs["base_high_color"],
        kwargs["base_mid_color"],
        kwargs["base_low_color"],
        kwargs["tubes"],
        kwargs["arrows"],
        kwargs["show_labels"],
        kwargs["vector_magnification"],
        kwargs["vector_width"],
        kwargs["blotches"],
        kwargs["blotch_size"],
        kwargs["show_out_of_tol"],
        kwargs["color_bar"],
        kwargs["show_percentages"],
        kwargs["show_fractions"],
        kwargs["x1"],
        kwargs["x2"],
        kwargs["high_tol"],
        kwargs["low_tol"],
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_point_coordinate(collection, group, pointname):
    """p489"""
    func_name = "Get Point Coordinate"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    Vector = NrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    log.debug(f"Vector: {Vector}")
    if Vector[0]:
        xVal = Vector[1]
        yVal = Vector[2]
        zVal = Vector[3]
        return (xVal, yVal, zVal)
    else:
        return False


def get_point_to_point_distance(p1col, p1grp, p1name, p2col, p2grp, p2name):
    """p495"""
    func_name = "Get Point To Point Distance"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("First Point", p1col, p1grp, p1name)
    NrkSdk.SetPointNameArg("Second Point", p2col, p2grp, p2name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    # xVal, yVal, zValNrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    xval = NrkSdk.GetDoubleArg("X Value", 0.0)
    yval = NrkSdk.GetDoubleArg("Y Value", 0.0)
    zval = NrkSdk.GetDoubleArg("Z Value", 0.0)
    mag = NrkSdk.GetDoubleArg("Magnitude", 0.0)
    returndict = {
        "xval": xval[1],
        "yval": yval[1],
        "zval": zval[1],
        "mag": mag[1],
    }
    return returndict


def set_default_colorization_options():
    """p531"""
    func_name = "Set Default Colorization Options"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColorizationOptionsArg(
        "Colorization Options",
        "Continuous",
        "Blue",
        "Green",
        "Red",
        False,
        True,
        False,
        100.0,
        1,
        False,
        0.1,
        False,
        False,
        True,
        False,
        0.5,
        -0.5,
        0.03,
        -0.5,
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_vector_group_display_attributes(magnification, blotch, tolerance):
    """p532"""
    func_name = "Set Vector Group Display Attributes"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetBoolArg("Draw Arrowheads?", False)
    NrkSdk.SetBoolArg("Indicate Values?", False)
    NrkSdk.SetDoubleArg("Vector Magnification", magnification)
    NrkSdk.SetIntegerArg("Vector Width", 1)
    NrkSdk.SetBoolArg("Draw Color Blotches?", False)
    NrkSdk.SetDoubleArg("Blotch Size", blotch)
    NrkSdk.SetBoolArg("Show Out of Tolerance Only?", False)
    NrkSdk.SetBoolArg("Show Color Bar in View?", True)
    NrkSdk.SetBoolArg("Show Color Bar Percentages?", False)
    NrkSdk.SetBoolArg("Show Color Bar Fractions?", False)
    # Available options:
    # "Deviation", "Sigma Rule", "Custom",
    NrkSdk.SetSaturationLimitTypeArg("High Saturation Limit Type", "Deviation")
    NrkSdk.SetDoubleArg("High Saturation Limit", 10.000000)
    # Available options:
    # "Deviation", "Sigma Rule", "Custom",
    NrkSdk.SetSaturationLimitTypeArg("Low Saturation Limit Type", "Deviation")
    NrkSdk.SetDoubleArg("Low Saturation Limit", -10.000000)
    NrkSdk.SetDoubleArg("High Tolerance", tolerance)
    NrkSdk.SetDoubleArg("Low Tolerance", (tolerance * -1))
    # Available options:
    # "Single Color", "Continuous", "Toleranced (Continuous)",
    # "Toleranced (Go / No-Go)", "Toleranced (Go / No-Go With Warning)", "Discrete Colors",
    NrkSdk.SetColorRangeMethodArg("Color Ranging Method", "Toleranced (Go / No-Go)")
    # Available options:
    # "Red", "Green", "Blue",
    NrkSdk.SetBaseColorTypeArg("Base High Color", "Blue")
    # Available options:
    # "Green", "Gray", "Red", "Blue",
    NrkSdk.SetBaseMidColorTypeArg("Base Mid Color", "Green")
    # Available options:
    # "Red", "Green", "Blue",
    NrkSdk.SetBaseColorTypeArg("Base Low Color", "Red")
    NrkSdk.SetBoolArg("Draw Tubes?", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def transform_object_by_delta_world_transform_operator(objects, transform):
    """p544"""
    func_name = "Transform Objects by Delta (World Transform Operator)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = []
    for item in objects:
        objNameList.append(f"{item[0]}::{item[1]}::Point Group")
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects to Transform", vObjectList)
    T = PythonToCSharp2Darray(transform)
    scale = 1.0
    vMatrixobj = System.Runtime.InteropServices.VariantWrapper(T)
    NrkSdk.SetWorldTransformArg("Delta Transform", vMatrixobj, scale)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def transform_objects_by_delta_about_working_frame(objects, transform):
    """p545"""
    func_name = "Transform Objects by Delta (About Working Frame)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = []
    for item in objects:
        objNameList.append(f"{item[0]}::{item[1]}::Point Group")
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects to Transform", vObjectList)
    T = PythonToCSharp2Darray(transform)
    vMatrixobj = System.Runtime.InteropServices.VariantWrapper(T)
    NrkSdk.SetTransformArg("Delta Transform", vMatrixobj)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def fit_geometry_to_point_group(
    geomType,
    dataCol,
    dataGroup,
    resultCol,
    resultName,
    profilename,
    reportdiv,
    fittol,
    outtol,
):
    """p547"""
    func_name = "Fit Geometry to Point Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    # Available options:
    #     "Line", "Plane", "Circle", "Sphere", "Cylinder",
    #     "Cone", "Paraboloid", "Ellipse", "Slot",
    NrkSdk.SetGeometryTypeArg("Geometry Type", geomType)
    NrkSdk.SetCollectionObjectNameArg("Group To Fit", dataCol, dataGroup)
    NrkSdk.SetCollectionObjectNameArg("Resulting Object Name", resultCol, resultName)
    NrkSdk.SetStringArg("Fit Profile Name", profilename)
    NrkSdk.SetBoolArg("Report Deviations", reportdiv)
    NrkSdk.SetDoubleArg("Fit Interface Tolerance (-1.0 use profile)", fittol)
    NrkSdk.SetBoolArg("Ignore Out of Tolerance Points", outtol)
    NrkSdk.SetCollectionObjectNameArg("Starting Condition Geometry (optional)", "", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def best_fit_group_to_group(
    refCollection,
    refGroup,
    corCollection,
    corGroup,
    showDialog,
    rmsTol,
    maxTol,
    allowScale,
):
    """p551"""
    func_name = "Best Fit Transformation - Group to Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Reference Group", refCollection, refGroup)
    NrkSdk.SetCollectionObjectNameArg("Corresponding Group", corCollection, corGroup)
    NrkSdk.SetBoolArg("Show Interface", showDialog)
    NrkSdk.SetDoubleArg("RMS Tolerance (0.0 for none)", rmsTol)
    NrkSdk.SetDoubleArg("Maximum Absolute Tolerance (0.0 for none)", maxTol)
    NrkSdk.SetBoolArg("Allow Scale", allowScale)
    NrkSdk.SetBoolArg("Allow X", True)
    NrkSdk.SetBoolArg("Allow Y", True)
    NrkSdk.SetBoolArg("Allow Z", True)
    NrkSdk.SetBoolArg("Allow Rx", True)
    NrkSdk.SetBoolArg("Allow Ry", True)
    NrkSdk.SetBoolArg("Allow Rz", True)
    NrkSdk.SetBoolArg("Lock Degrees of Freedom", False)
    NrkSdk.SetBoolArg("Generate Event", False)
    NrkSdk.SetFilePathArg("File Path for CSV Text Report (requires Show Interface = TRUE)", "", False)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    results = {}

    T = System.Runtime.InteropServices.VariantWrapper([])
    trans_in_work = NrkSdk.GetTransformArg("Transform in Working", T)
    results["trans_in_work"] = CSharpToPython2Darray(trans_in_work[1])

    T = System.Runtime.InteropServices.VariantWrapper([])
    trans_optimum = NrkSdk.GetWorldTransformArg("Optimum Transform", T, 0.0)
    results["trans_in_world"] = CSharpToPython2Darray(trans_optimum[1])
    results["scale"] = trans_optimum[2]

    results["rms"] = NrkSdk.GetDoubleArg("RMS Deviation", 0.0)[1]
    results["max_abs_dev"] = NrkSdk.GetDoubleArg("Maximum Absolute Deviation", 0.0)[1]
    results["n_unknown"] = NrkSdk.GetIntegerArg("Number of Unknowns", 0.0)[1]
    results["n_equations"] = NrkSdk.GetIntegerArg("Numnber of Equations", 0.0)[1]
    results["robustness"] = NrkSdk.GetDoubleArg("Robustness", 0.0)[1]

    return results


def get_measurement_weather_data(collection, group, pointname):
    """p557"""
    func_name = "Get Measurement Weather Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    temp = NrkSdk.GetDoubleArg("Temperature (deg F)", 0.0)
    pres = NrkSdk.GetDoubleArg("Pressure (in. Hg)", 0.0)
    humi = NrkSdk.GetDoubleArg("Humidity (% RH)", 0.0)
    returndict = {
        "temperature": (FahrenheitToCelsius(temp[1]), "\xb0C"),
        "pressure": (InchHgtoMilliBar(pres[1]), "mBar"),
        "humidity": (humi[1], "%RH"),
        "sTemperature": ("{:.1f}".format(FahrenheitToCelsius(temp[1])), "\xb0C"),
        "sPressure": ("{:.1f}".format(InchHgtoMilliBar(pres[1])), "mBar"),
        "sHumidity": ("{:.1f}".format(humi[1]), "%RH"),
    }
    return returndict


def get_measurement_auxiliary_data(collection, group, pointname, aux):
    """p558"""
    func_name = "Get Measurement Auxiliary Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.SetStringArg("Auxiliary Name", aux)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    value = NrkSdk.GetDoubleArg("Value", 0.0)
    units = NrkSdk.GetStringArg("Units", "")
    returndict = {
        "value": value[1],
        "units": units[1],
    }
    return returndict


def get_measurement_info_data(collection, group, pointname):
    """p559"""
    func_name = "Get Measurement Info Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    value = NrkSdk.GetStringArg("Info Data", "")
    if not value[1]:
        return False

    # Clean the results before returning them
    results = []
    values = value[1].split(";")
    for val in values:
        results.append(val.strip())

    return results


def make_point_to_point_relationship(relCol, relName, p1col, p1grp, p1, p2col, p2grp, p2):
    """p632"""
    func_name = "Make Point to Point Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetPointNameArg("First Point Name", p1col, p1grp, p1)
    NrkSdk.SetPointNameArg("Second Point Name", p2col, p2grp, p2)
    NrkSdk.SetToleranceVectorOptionsArg(
        "Tolerance",
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
    )
    NrkSdk.SetToleranceVectorOptionsArg(
        "Constraint",
        True,
        0.0,
        True,
        0.0,
        True,
        0.0,
        False,
        0.0,
        True,
        0.0,
        True,
        0.0,
        True,
        0.0,
        False,
        0.0,
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


def make_group_to_nominal_group_relationship(
    relCol,
    relName,
    nomCol,
    nomGrp,
    meaCol,
    meaGrp,
):
    """p646"""
    func_name = "Make Group to Nominal Group Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetCollectionObjectNameArg("Nominal Group Name", nomCol, nomGrp)
    NrkSdk.SetCollectionObjectNameArg("Measured Group Name", meaCol, meaGrp)
    NrkSdk.SetBoolArg("Auto Update a Vector Group?", False)
    NrkSdk.SetBoolArg("Use Closest Point?", True)
    NrkSdk.SetBoolArg("Display Closest Point Watch Window?", False)
    NrkSdk.SetBoolArg("Use View Zooming With Proximity?", False)
    NrkSdk.SetBoolArg("Ignore Points Beyond Threshold?", False)
    NrkSdk.SetDoubleArg("Proximity Threshold?", 0.01)
    NrkSdk.SetToleranceVectorOptionsArg(
        "Tolerance",
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        True,
        1.5,
        False,
        0.0,
        False,
        0.0,
        False,
        0.0,
        True,
        0.0,
    )
    NrkSdk.SetToleranceVectorOptionsArg(
        "Constraint",
        True,
        0.0,
        True,
        0.0,
        True,
        0.0,
        False,
        0.0,
        True,
        0.0,
        True,
        0.0,
        True,
        0.0,
        False,
        0.0,
    )
    NrkSdk.SetDoubleArg("Fit Weight", 1.0)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def make_geometry_fit_and_compare_to_nominal_relationship(
    relCol, relName, nomCol, nomName, data, resultCol, resultName
):
    """p649"""
    func_name = "Make Geometry Fit and Compare to Nominal Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetCollectionObjectNameArg("Nominal Geometry", nomCol, nomName)
    objNameList = []
    for item in data:
        objNameList.append(f"{item[0]}::{item[1]}::Point Group")
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Point Groups to Fit", vObjectList)
    NrkSdk.SetCollectionObjectNameArg("Resulting Object Name (Optional)", resultCol, resultName)
    NrkSdk.SetStringArg("Fit Profile Name (Optional)", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_relationship(collection, relationship):
    """p656"""
    func_name = "Delete Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationship)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_general_relationship_statistics(collection, relationshipname):
    """p659"""
    func_name = "Get General Relationship Statistics"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationshipname)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    results = {}
    results["max_dev"] = NrkSdk.GetDoubleArg("Max Deviation", 0.0)[1]
    results["rms"] = NrkSdk.GetDoubleArg("RMS", 0.0)[1]
    results["has_sign_dev"] = NrkSdk.GetBoolArg("Has Signed Deviation?", False)[1]
    results["sign_max_dev"] = NrkSdk.GetDoubleArg("Signed Max Deviation", 0.0)[1]
    results["sign_min_dev"] = NrkSdk.GetDoubleArg("Signed Min Deviation", 0.0)[1]
    return results


def get_geom_relationship_criteria(relCol, relName, criteria):
    """p660"""
    func_name = "Get Geom Relationship Criteria"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetStringArg("Criteria", criteria)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    results = {}
    # For each 'Get___' request a tuple is returned, we need to save the second argument ([1]) only.
    results["nominal"] = NrkSdk.GetDoubleArg("Nominal", 0.0)[1]
    results["measured"] = NrkSdk.GetDoubleArg("Measured", 0.0)[1]
    results["delta"] = NrkSdk.GetDoubleArg("Delta", 0.0)[1]
    results["lowtol"] = NrkSdk.GetDoubleArg("Low Tolerance", 0.0)[1]
    results["hightol"] = NrkSdk.GetDoubleArg("High Tolerance", 0.0)[1]
    results["deltaweight"] = NrkSdk.GetDoubleArg("Optimization: Delta Weight", 0.0)[1]
    results["outoftolweight"] = NrkSdk.GetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)[1]
    return results


def set_relationship_associated_data(relCol, group, collectionmeasured, groupmeasured):
    """p664"""
    # The function only excepts 'groups' as an input.
    # Individual points, point clouds and objects aren't support for now.
    func_name = "Set Relationship Associated Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, group)
    # individual points
    ptNameList = []
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg("Individual Points", vPointObjectList)
    # point groups
    objNameList = [
        f"{collectionmeasured}::{groupmeasured}::Point Group",
    ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Point Groups", vObjectList)
    # point cloud
    objNameList = []
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Point Clouds", vObjectList)
    # objects
    objNameList = []
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects", vObjectList)
    # additional setting
    NrkSdk.SetBoolArg("Ignore Empty Arguments?", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_reporting_frame(relCol, relGrp, frmCol, frmName):
    """p679"""
    func_name = "Set Relationship Reporting Frame"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    NrkSdk.SetCollectionObjectNameArg("Reporting Frame", frmCol, frmName)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_geom_relationship_criteria(relCol, relName, criteriatype):
    """p680"""
    func_name = "Set Geom Relationship Criteria"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    # Flatness
    if criteriatype == "Flatness":
        NrkSdk.SetStringArg("Criteria", "Flatness")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.1, True, -0.1)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Centroid Z
    elif criteriatype == "Centroid Z":
        NrkSdk.SetStringArg("Criteria", "Centroid Z")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Avg Distance Between
    elif criteriatype == "Avg Dist Between":
        NrkSdk.SetStringArg("Criteria", "Avg Dist Between")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Length
    elif criteriatype == "Length":
        NrkSdk.SetStringArg("Criteria", "Length")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.05, True, -0.05)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    else:
        log.warning("Incorrect criteria type set!")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_geom_relationship_cardinal_points(collection, relationship, groupname):
    """p688"""
    func_name = "Set Geom Relationship Cardinal Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationship)
    NrkSdk.SetBoolArg("Create Cardinal Pts when Fitting?", True)
    NrkSdk.SetBoolArg("Prefix Cardinal Pts name with Rel name?", True)
    NrkSdk.SetStringArg("Cardinal Pts Group Name", groupname)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_geom_relationship_cardinal_points(collection, relationship):
    """p689"""
    func_name = "Get Geom Relationship Cardinal Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationship)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False

    ptNameList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Cardinal Point Name List", ptNameList)
    if objectList[0]:
        points = []
        for i in range(objectList[1].GetLength(0)):
            points.append(NamedPoint(objectList[1][i].split("::")))
        return points
    else:
        return False


def set_geom_relationship_auto_vectors_nominal_avn(collection, relationshipname, bool):
    """p690"""
    func_name = "Set Geom Relationship Auto Vectors Nominal (AVN)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationshipname)
    NrkSdk.SetBoolArg("Create Auto Vectors AVN", bool)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_auto_vectors_fit_avf(collection, relationshipname, bool):
    """p691"""
    func_name = "Set Relationship Auto Vectors Fit (AVF)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationshipname)
    NrkSdk.SetBoolArg("Create Auto Vectors AVF", bool)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_desired_meas_count(relCol, relName, count):
    """p693"""
    func_name = "Set Relationship Desired Meas Count"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetIntegerArg("Desired Measurement Count", count)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_tolerance_vector_type(relCol, relName):
    """p697"""
    func_name = "Set Relationship Tolerance (Vector Type)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetToleranceVectorOptionsArg(
        "Vector Tolerance",
        False,
        0.000000,
        False,
        0.000000,
        False,
        0.000000,
        True,
        1.0,
        False,
        0.000000,
        False,
        0.000000,
        False,
        0.000000,
        True,
        0.0,
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ###################################
# Chapter 9 - Reporting Operations ##
# ###################################
def set_vector_group_report_options(collection, vector_group, **kwargs):
    """p769"""
    func_name = "Set Vector Group Report Options"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Vector Group", collection, vector_group)
    NrkSdk.SetPointDeltaReportOptionsArg(
        "Report Options",
        kwargs["coord_system"],  # str
        kwargs["record_format"],  # str
        kwargs["point_a"],  # bool
        kwargs["point_b"],  # bool
        kwargs["delta"],  # bool
        kwargs["mag"],  # bool
        kwargs["X"],  # bool
        kwargs["Y"],  # bool
        kwargs["Z"],  # bool
        kwargs["sort_point_names"],  # bool
        kwargs["tolerance_values"],  # bool
        kwargs["colorize"],  # bool
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_report_options(relCol, relGrp, **kwargs):
    """p770"""
    func_name = "Set Relationship Report Options"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    NrkSdk.SetPointDeltaReportOptionsArg(
        "Report Options",
        kwargs["coord_system"],
        kwargs["record_format"],
        kwargs["a"],
        kwargs["b"],
        kwargs["c"],
        kwargs["d"],
        kwargs["e"],
        kwargs["f"],
        kwargs["summary"],
        kwargs["h"],
        kwargs["i"],
        kwargs["j"],
    )
    NrkSdk.ExecuteStep()
    getResult(func_name)


def notify_user_text_array(txt, timeout=0):
    """p804"""
    func_name = "Notify User Text Array"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    stringList = [
        txt,
    ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Notification Text", vStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.SetBoolArg("Auto expand to fit text?", False)
    NrkSdk.SetIntegerArg("Display Timeout", timeout)
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ####################################
# Chapter 10 - Excel Direct Connect ##
# ####################################


# ##############################################
# Chapter 11 - MS Office Reporting Operations ##
# ##############################################


# #####################################
# Chapter 12 - Instrument Operations ##
# #####################################
def get_last_instrument_index(collection):
    """p870"""
    func_name = "Get Last Instrument Index"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    ColInstID = NrkSdk.GetColInstIdArg("Instrument ID", "", 0)
    collection = ColInstID[1]
    inst_id = ColInstID[2]
    log.debug(f"ColInstID: {collection}, {inst_id}")
    return (collection, inst_id)


def point_at_target(instCol, instId, collection, group, name):
    """p877"""
    func_name = "Point At Target"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", instCol, instId)
    NrkSdk.SetPointNameArg("Target ID", collection, group, name)
    NrkSdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        raise SystemError(f"Failed pointing at point: {collection}::{group}::{name}")


def add_new_instrument(instName):
    """p889"""
    func_name = "Add New Instrument"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetInstTypeNameArg("Instrument Type", instName)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    ColInstID = NrkSdk.GetColInstIdArg("Instrument Added (result)", "", 0)
    log.debug(f"ColInstID: {ColInstID}")
    return ColInstID[1], ColInstID[2]


def start_instrument(InstCol, InstID, initialize, simulation):
    """p902"""
    func_name = "Start Instrument Interface"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", InstCol, InstID)
    NrkSdk.SetBoolArg("Initialize at Startup", initialize)
    NrkSdk.SetStringArg("Device IP Address (optional)", "")
    NrkSdk.SetIntegerArg("Interface Type (0=default)", 0)
    NrkSdk.SetBoolArg("Run in Simulation", simulation)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def stop_instrument(collection, inst_id):
    """p903"""
    func_name = "Stop Instrument Interface"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection, inst_id)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def verify_instrument_connection(collection, inst_id):
    """p905"""
    func_name = "Verify Instrument Connection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection, inst_id)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    bValue = NrkSdk.GetBoolArg("Connected?", False)
    log.debug(f"Tracker Connection info: {bValue[1]}")
    if bValue[0]:
        return bValue[1]
    else:
        return False


def configure_and_measure(
    instrumentCollection,
    instId,
    targetCollection,
    group,
    name,
    profileName,
    measureImmediately,
    waitForCompletion,
    timeoutInSecs,
):
    """p906"""
    func_name = "Configure and Measure"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", instrumentCollection, instId)
    NrkSdk.SetPointNameArg("Target Name", targetCollection, group, name)
    NrkSdk.SetStringArg("Measurement Mode", profileName)
    NrkSdk.SetBoolArg("Measure Immediately", measureImmediately)
    NrkSdk.SetBoolArg("Wait for Completion", waitForCompletion)
    NrkSdk.SetDoubleArg("Timeout in Seconds", timeoutInSecs)
    NrkSdk.ExecuteStep()
    result = getResult(func_name)
    return result


def compute_CTE_scale_factor(cte, parttemp):
    """p929"""
    func_name = "Compute CTE Scale Factor"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    NrkSdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    NrkSdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    scaleFactor = NrkSdk.GetDoubleArg("Scale Factor", 0.0)
    if scaleFactor[0]:
        return scaleFactor[1]
    else:
        return False


def set_instrument_scale_absolute(collection, instid, scaleFactor):
    """p931"""
    func_name = "Set (absolute) Instrument Scale Factor (CAUTION!)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection, instid)
    NrkSdk.SetDoubleArg("Scale Factor", scaleFactor)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_observation_info(collection, group, pointname, index=0):
    """p951"""
    func_name = "Get Observation Info"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.SetIntegerArg("Observation Index", index)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False

    inst = NrkSdk.GetColInstIdArg("Resulting Instrument", "", 0)
    vector = NrkSdk.GetVectorArg("Resultant Vector", 0.0, 0.0, 0.0)
    active = NrkSdk.GetBoolArg("Active?", False)
    timestamp = NrkSdk.GetStringArg("Timestamp", "")
    rmsError = NrkSdk.GetDoubleArg("RMS Error", 0.0)
    temperature = NrkSdk.GetDoubleArg("Temperature (deg F)", 0.0)
    pressure = NrkSdk.GetDoubleArg("Pressure (in. Hg)", 0.0)
    humidity = NrkSdk.GetDoubleArg("Humidity (% RH)", 0.0)
    infoData = NrkSdk.GetStringArg("Info Data", "")
    results = {
        "instCol": inst[1],
        "instId": inst[2],
        "vec_xVal": vector[1],
        "vec_yVal": vector[2],
        "vec_zVal": vector[3],
        "active": active[1],
        "timestamp": timestamp[1],
        "rmsError": rmsError[1],
        "temperature": temperature[1],
        "pressure": pressure[1],
        "humidity": humidity[1],
        "infoData": infoData[1],
    }
    return results


# ################################
# Chapter 13 - Robot Operations ##
# ################################


# ##################################
# Chapter 14 - Utility Operations ##
# ##################################
def delete_folder(foldername):
    """p1054"""
    func_name = "Delete Folder"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Folder Path", foldername)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def move_collection_to_folder(collection, folder):
    """p1055"""
    func_name = "Move Collection to Folder"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_folders_by_wildcard(search):
    """p1057"""
    func_name = "Get Folders by Wildcard"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Search String", search)
    NrkSdk.SetBoolArg("Case Sensitive Search", False)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    stringList = System.Runtime.InteropServices.VariantWrapper([])
    vStringList = NrkSdk.GetStringRefListArg("Folder List", stringList)
    if vStringList[0]:
        folders = []
        for i in range(vStringList[1].GetLength(0)):
            folders.append(vStringList[1][i])
        return folders
    else:
        return False


def get_folder_collections(folder):
    """p1060"""
    func_name = "Get Folder Collections"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    stringList = System.Runtime.InteropServices.VariantWrapper([])
    vStringList = NrkSdk.GetStringRefListArg("Collection List", stringList)
    if vStringList[0]:
        folders = []
        for i in range(vStringList[1].GetLength(0)):
            folders.append(vStringList[1][i])
        return folders
    else:
        return False


def set_collection_notes(collection, notes):
    """p1063"""
    func_name = "Set Collection Notes"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    stringList = [notes]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Notes", vStringList)
    NrkSdk.SetBoolArg("Append? (FALSE = Overwrite)", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_working_frame(col, name):
    """p1080"""
    func_name = "Set Working Frame"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("New Working Frame Name", col, name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_objects(col, name, objtype):
    """p1089"""
    func_name = "Delete Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = [
        f"{col}::{name}::{objtype}",
    ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Object Names", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_interaction_mode(sa_interaction_mode, mp_interaction_mode, mp_dialog_interaction_mode):
    """p1117"""
    func_name = "Set Interaction Mode"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    # Available options:
    # "Manual", "Automatic", "Silent",
    NrkSdk.SetSAInteractionModeArg("SA Interaction Mode", sa_interaction_mode)
    # Available options:
    # "Halt on Failure Only", "Halt on Failure or Partial Success", "Never Halt",
    NrkSdk.SetMPInteractionModeArg("Measurement Plan Interaction Mode", mp_interaction_mode)
    # Available options:
    # "Block Application Interaction", "Allow Application Interaction",
    NrkSdk.SetMPDialogInteractionModeArg("Measurement Plan Dialog Interaction Mode", mp_dialog_interaction_mode)
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ###########################################
# Chapter 15 - Accumulator Math Operations ##
# ###########################################


# ######################################
# Chapter 16 - Scalar Math Operations ##
# ######################################


# ######################################
# Chapter 17 - Vector Math Operations ##
# ######################################


# ###############################
# Chapter 18 - RMP Subroutines ##
# ###############################


# #########################
# Chapter 19 - Variables ##
# #########################
