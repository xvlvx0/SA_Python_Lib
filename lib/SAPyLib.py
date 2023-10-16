# -*- coding: utf-8 -*-
"""
SAPyLib = Spatial Analyzer Python .NET library

This library provides the communication layer to the SA SDK .NET dll.
Author: L. Ververgaard
"""
import sys
import os
import logging

# Get the logger
log = logging.getLogger(__name__)


# This library was developed on Python 3.9, but due to pythonnet requirements (max version available is Python3.8) it needed some back porting.
if sys.version_info.major == 2:
    raise OSError("This version of Python isn't supported. Version 3.8 is minimum.")
elif sys.version_info.major == 3 and sys.version_info.minor <= 7:
    raise OSError("This version of Python isn't supported. Version 3.8 is minimum.")
elif sys.version_info.major == 3 and sys.version_info.minor == 8:
    from typing import Union
    from typing import Tuple as tuple
    from typing import List as list
elif sys.version_info.major == 3 and sys.version_info.minor >= 9:
    from typing import Union

import clr  # #                          Python.NET library

clr.AddReference("System")  # #             Import via python.net the .NET System Library
clr.AddReference("System.Collections")  # # Import via python.net the .NET System.Collections Library
clr.AddReference("System.Reflection")  # #  Import via python.net the .NET System.Reflection Library
import System
from System import Array, Double, String
from System.Collections.Generic import List
import System.Reflection


BASE_PATH = r"C:\Analyzer Data\Scripts\SA_Python_Lib"
DLL_FOLDER = os.path.join(BASE_PATH, "dll")


# Get SA SDK dll
sa_sdk_dll_file = os.path.join(DLL_FOLDER, "Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
NrkSdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = NrkSdk.Connect("127.0.0.1")
if not SAConnected:
    raise IOError("Connection to SA SDK failed!")


# Get SA Python Tools dll
sa_py_tools_dll_file = os.path.join(DLL_FOLDER, "SA_Python_Tools.dll")
sa_py_tools_dll = System.Reflection.Assembly.LoadFile(sa_py_tools_dll_file)
sa_py_tools_class_type = sa_py_tools_dll.GetType("SA_Python_Tools.SA_Py_Tool")
sa_py_tools = System.Activator.CreateInstance(sa_py_tools_class_type)


# Get the logger
log = logging.getLogger(__name__)


# ###############
# Base methods ##
# ###############
def getResult_Bare(func_name: str) -> tuple[str, str]:
    """Get the methods execution result without processing."""
    boolean, result = NrkSdk.GetMPStepResult(0)
    log.debug(f"getResult_Bare --> {func_name}: {boolean}, {result}")
    return (boolean, result)


def getResult(func_name: str) -> bool:
    """Get the methods execution result and process the result."""
    boolean, result = NrkSdk.GetMPStepResult(0)
    if result == -1:
        # SDKERROR = -1
        log.error(f"{func_name}: {boolean}, {result}")
        MPStepMessages()
        raise SystemError("Execution raised: SDKERROR!")
    elif result == 0:
        # UNDONE = 0
        log.warning(f"getResult --> {func_name}: {boolean}, {result}")
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
        # DONE SUCCESS = 2
        return True
    elif result == 3:
        # DONE FATAL ERROR = 3
        log.error(f"getResult --> ERROR: {func_name}: {boolean}, {result}")
        log.error("Execution: FAILED!")
        MPStepMessages()
        return False
    elif result == 4:
        # DONE MINOR ERROR = 4
        log.warning(f"getResult --> {func_name}: {boolean}, {result}")
        log.warning("Execution: FAILED - minor error!")
        MPStepMessages()
        return False
    elif result == 5:
        # CURRENT TASK = 5
        log.debug(f"getResult --> {func_name}: {boolean}, {result}")
        log.info("Execution: current task")
        MPStepMessages()
        return True
    elif result == 6:
        # UNKNOWN = 6
        log.debug(f"getResult --> {func_name}: {boolean}, {result}")
        log.info("I have no clue!")
        MPStepMessages()
        raise SystemError("I have no clue!")
    else:
        return False


def MPStepMessages() -> None:
    """Get the MPStep messages."""
    log.debug("Get the MPStep messages.")
    stringList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    try:
        vStringList = NrkSdk.GetMPStepMessages(stringList)
        if vStringList[0]:
            for i in range(vStringList[1].GetLength(0)):
                log.info(f"MPStepMessage: {vStringList[1][i]}")
        else:
            log.info("MPStepMessage: 'NO MESSAGES'")
    except System.Runtime.InteropServices.COMException as err:
        log.error(f"Getting MP Step failed with error: {err}")


class Point3D:
    def __init__(self, x: float, y: float, z: float) -> None:
        self.X = x
        self.Y = y
        self.Z = z


class NamedPoint:
    def __init__(self, point: list[str]) -> None:
        if len(point) != 3:
            raise ValueError(f"Named Point needs 3 arguments. You provided: {len(point)} arguments.")

        self.collection = point[0]
        self.group = point[1]
        self.name = point[2]


class NamedPoint3D:
    def __init__(self, point: Union[NamedPoint, str, list[str]], xyz: list[float] = []) -> None:
        log.debug("NamedPoint3D")

        if isinstance(point, NamedPoint):
            self.collection = point.collection
            self.group = point.group
            self.name = point.name
        else:
            if isinstance(point, str):
                point = point.split("::")

            if isinstance(point, list):
                if len(point) == 3:
                    self.collection = point[0]
                    self.group = point[1]
                    self.name = point[2]
                elif len(point) == 1:
                    self.collection = ""
                    self.group = ""
                    self.name = point[0]
                else:
                    raise ValueError("Point name is in an incorrect format!")

        if xyz:
            self.X = xyz[0]
            self.Y = xyz[1]
            self.Z = xyz[2]
        else:
            # Get XYZ from SA
            point_obj = get_point_coordinate(self.collection, self.group, self.name)
            log.info(f"pData: {point_obj}")
            self.X = point_obj.X
            self.Y = point_obj.Y
            self.Z = point_obj.Z


def FahrenheitToCelsius(tempF: float) -> float:
    """Convert Fahrenheit to Celsius."""
    tempC = (tempF - 32) * (5 / 9)
    return tempC


def InchHgtoMilliBar(pressInchHg: float) -> float:
    """Convert Inch Mercury (Hg) to mBar."""
    pressMilliBar = pressInchHg * 33.864
    return pressMilliBar


def python_list_to_csharp_2D_array(input_list: list, array_depth: tuple = (4, 4)) -> Array:
    """Convert a python N-List to a C# Array."""
    CSharpArray = Array.CreateInstance(Double, array_depth[0], array_depth[1])
    for i in range(array_depth[0]):
        for j in range(array_depth[1]):
            CSharpArray.SetValue(input_list[i][j], i, j)
    return CSharpArray


def python_list_to_csharp_list(input_list: list) -> List:
    output_list = List[String]()
    for item in input_list:
        output_list.Add(item)
    return output_list


def csharp_array_to_python_2D_list(input_array: Array, array_depth: tuple = (4, 4)) -> list:
    """Convert a C# Array to a N-List."""
    python_list = []
    python_list.append([])
    for i in range(array_depth[0]):
        for j in range(array_depth[1]):
            python_list[i].append(input_array.GetValue(i, j))
        python_list.append([])
    return python_list


# ##############################
# Chapter 2 - File Operations ##
# ##############################
def find_files_in_directory(directory: str, searchPattern: str) -> list:
    """p29"""
    func_name = "Find Files in Directory"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Directory", directory)
    NrkSdk.SetStringArg("File Name Pattern", searchPattern)
    NrkSdk.SetBoolArg("Recursive?", False)
    NrkSdk.ExecuteStep()
    results = getResult(func_name)
    if not results:
        return []
    stringList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    vStringList = NrkSdk.GetStringRefListArg("Files", stringList)
    if vStringList[0]:
        files = []
        for i in range(vStringList[1].GetLength(0)):
            files.append(vStringList[1][i])
        return files
    else:
        return []


# ######################################
# Chapter 3 - Process Flow Operations ##
# ######################################
def ask_for_string(question: str, initialanswer: str = "") -> str:
    """p123"""
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


def ask_for_string_pulldown(question: str, answers: list) -> str:
    """p124"""
    func_name = "Ask for String (Pull-Down Version)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    # question section
    question_list = python_list_to_csharp_list([question])
    QvStringList = sa_py_tools.GetListWrapper(question_list)
    NrkSdk.SetStringRefListArg("Question or Statement", QvStringList)
    # answers section
    answers_list = python_list_to_csharp_list(answers)
    AvStringList = sa_py_tools.GetListWrapper(answers_list)
    NrkSdk.SetStringRefListArg("Possible Answers", AvStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 12, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    answer = NrkSdk.GetStringArg("Answer", "")
    log.debug(f"Answer: {answer}")
    return answer[1]


# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################


# ###########################
# Chapter 5 - View Control ##
# ###########################
def show_objects(collection: str, objects: str, name: str) -> None:
    """p161"""
    func_name = "Show Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = python_list_to_csharp_list([f"{collection}::{objects}::{name}"])
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Show", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def hide_objects(collection: str, name: str, objtype: str) -> None:
    """p163"""
    func_name = "Hide Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = python_list_to_csharp_list([f"{collection}::{name}::{objtype}"])
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Hide", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def show_hide_by_object_type(collection: str, objtype: str, hide: bool) -> None:
    """p164"""
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

    # Hide: hide=True
    # Show: hide=False
    log.debug(f"{collection}::{objtype} Hide? {hide}")
    NrkSdk.SetBoolArg("Hide? (Show = FALSE)", hide)

    NrkSdk.ExecuteStep()
    getResult(func_name)


def show_hide_callout_view(collection: str, calloutname: str, show: bool) -> None:
    """p167"""
    func_name = "Show / Hide Callout View"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Callout View To Show", collection, calloutname)

    # Show: show=True
    # Hide: show=False
    NrkSdk.SetBoolArg("Show Callout View?", show)

    NrkSdk.ExecuteStep()
    getResult(func_name)


def hide_all_callout_views() -> None:
    """p168"""
    func_name = "Hide All Callout Views"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def center_graphics_about_objects(objtype: str = "Any", ColWild: str = "*", ObjWild: str = "*") -> None:
    """p198"""
    func_name = "Center Graphics About Object(s)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    # Available options:
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse",
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter",
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh",
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group",
    NrkSdk.SetObjectTypeArg("Object Type", objtype)

    NrkSdk.SetStringArg("Collection Wildcard Criteria", ColWild)
    NrkSdk.SetStringArg("Object Wildcard Criteria", ObjWild)
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ######################################
# Chapter 6 - Cloud Viewer Operations ##
# ######################################


# ######################################
# Chapter 7 - Construction Operations ##
# ######################################
def rename_point(
    orgCol: str, orgGrp: str, orgName: str, newCol: str, newGrp: str, newName: str, overwrite: bool = False
) -> None:
    """p216"""
    func_name = "Rename Point"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Original Point Name", orgCol, orgGrp, orgName)
    NrkSdk.SetPointNameArg("New Point Name", newCol, newGrp, newName)
    NrkSdk.SetBoolArg("Overwrite if exists?", overwrite)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        raise SystemError(f"Renaming point: '{orgCol}::{orgGrp}::{orgName}' failed.")


def rename_collection(fromName: str, toName: str) -> None:
    """p218"""
    func_name = "Rename Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Original Collection Name", fromName)
    NrkSdk.SetCollectionNameArg("New Collection Name", toName)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        raise SystemError(f"Renaming folder: '{fromName}' failed!")


def rename_object(old_col: str, old_name: str, new_col: str, new_name: str) -> None:
    """p219"""
    func_name = "Rename Object"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Original Object Name", old_col, old_name)
    NrkSdk.SetCollectionObjectNameArg("New Object Name", new_col, new_name)
    NrkSdk.SetBoolArg("Overwrite if exists?", False)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        # raise SystemError(f"Renaming object: '{old_col}::{old_name}' failed!")
        log.error(f"Renaming object: '{old_col}::{old_name}' failed!")


def delete_points(collection: str, group: str, name: str) -> bool:
    """p221"""
    func_name = "Delete Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    ptNameList = python_list_to_csharp_list([f"{collection}::{group}::{name}"])
    vPointObjectList = sa_py_tools.GetListWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg("Point Names", vPointObjectList)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        log.error("Failed to delete points")
        return False
    return True


def delete_points_wildcard_selection(collection: str, group: str, name: str, objtype: str) -> None:
    """p222"""
    func_name = "Delete Points WildCard Selection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objNameList = python_list_to_csharp_list([f"{collection}::{group}::{objtype}"])
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Groups to Delete From", vObjectList)
    NrkSdk.SetPointNameArg("WildCard Selection Names", "*", "*", name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_objects_from_surface_faces_runtime_select(facetype: str = "") -> None:
    """p223"""
    func_name = "Construct Objects From Surface Faces - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

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


def set_or_construct_default_collection(collection: str) -> None:
    """p225"""
    func_name = "Set (or construct) default collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection Name", collection)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_collection(collection: str, make_default: bool = True) -> None:
    """p226"""
    func_name = "Construct Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection Name", collection)
    NrkSdk.SetStringArg("Folder Path", "")

    # default collection: makeDefault=True
    # not default collection: makeDefault=False
    NrkSdk.SetBoolArg("Make Default Collection?", make_default)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_active_collection_name() -> str:
    """p203"""
    func_name = "Get Active Collection Name"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetCollectionNameArg("Currently Active Collection Name", "")
    if not sValue[0]:
        return ""
    return sValue[1]


def delete_collection(collection: str) -> None:
    """p228"""
    func_name = "Delete Collection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_a_point_in_working_coordinates(
    collection: str, group: str, name: str, x: float, y: float, z: float
) -> None:
    """p232"""
    func_name = "Construct a Point in Working Coordinates"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.SetVectorArg("Working Coordinates", x, y, z)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_point_at_intersection_of_plane_and_line(
    collection_plane: str,
    name_plane: str,
    collection_line: str,
    name_line: str,
    collection_point: str,
    group_point: str,
    name_point: str,
) -> None:
    """p243"""
    func_name = "Construct Point at Intersection of Plane and Line"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", collection_plane, name_plane)
    NrkSdk.SetCollectionObjectNameArg("Line Name", collection_line, name_line)
    NrkSdk.SetPointNameArg("Resulting Point Name", collection_point, group_point, name_point)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_line_2_points(
    collection_line: str,
    name_line: str,
    collection_1: str,
    group_1: str,
    name_1: str,
    collection_2: str,
    group_2: str,
    name_2: str,
) -> None:
    """p290"""
    func_name = "Construct Line 2 Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Line Name", collection_line, name_line)
    NrkSdk.SetPointNameArg("First Point", collection_1, group_1, name_1)
    NrkSdk.SetPointNameArg("Second Point", collection_2, group_2, name_2)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_plane(collection_plane: str, name_plane: str) -> None:
    """p300"""
    func_name = "Construct Plane"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", collection_plane, name_plane)
    NrkSdk.SetVectorArg("Plane Center (in working coordinates)", 0.0, 0.0, 0.0)
    NrkSdk.SetVectorArg("Plane Normal (in working coordinates)", 0.0, 0.0, 1.0)
    NrkSdk.SetDoubleArg("Plane Edge Dimension", 0.0)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def construct_frame_known_origin_object_direction_object_direction(
    collection_point: str,
    group_point: str,
    name_point: str,
    x: float,
    y: float,
    z: float,
    collection_pri_ax: str,
    name_obj_pri_ax: str,
    name_pri_ax: str,
    collection_sec_ax: str,
    name_obj_sec_ax: str,
    name_sec_ax: str,
    collection_frame: str,
    name_frame: str,
) -> None:
    """p358"""
    func_name = "Construct Frame, Known Origin, Object Direction, Object Direction"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Known Point", collection_point, group_point, name_point)
    NrkSdk.SetVectorArg("Known Point Value in New Frame", x, y, z)
    NrkSdk.SetCollectionObjectNameArg("Primary Axis Object", collection_pri_ax, name_obj_pri_ax)

    # Available options:
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis", "-Z Axis",
    NrkSdk.SetAxisNameArg("Primary Axis Defines Which Axis", name_pri_ax)

    NrkSdk.SetCollectionObjectNameArg("Secondary Axis Object", collection_sec_ax, name_obj_sec_ax)

    # Available options:
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis", "-Z Axis",
    NrkSdk.SetAxisNameArg("Secondary Axis Defines Which Axis", name_sec_ax)

    NrkSdk.SetCollectionObjectNameArg("Frame Name (Optional)", collection_frame, name_frame)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def create_relationship_callout(
    collection_callout: str,
    name_callout: str,
    collection_relationship: str,
    name_relationship: str,
    xpos: float = 0.1,
    ypos: float = 0.1,
) -> None:
    """p400"""
    func_name = "Create Relationship Callout"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection_callout, name_callout)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    vStringList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    NrkSdk.SetEditTextArg("Additional Notes (blank for none)", vStringList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def create_text_callout(
    collection_callout: str, name_callout: str, text: str, xpos: float = 0.1, ypos: float = 0.1
) -> None:
    """p402"""
    func_name = "Create Text Callout"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection_callout, name_callout)
    stringList = python_list_to_csharp_list([text])
    vStringList = sa_py_tools.GetListWrapper(stringList)
    NrkSdk.SetEditTextArg("Text", vStringList)
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    NrkSdk.SetPointNameArg("Callout Anchor Point (Optional)", "", "", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_default_callout_view_properties(name_callout: str) -> None:
    """p409"""
    func_name = "Set Default Callout View Properties"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Default Callout View Name", name_callout)
    NrkSdk.SetBoolArg("Lock View Point?", True)
    NrkSdk.SetBoolArg("Recall Working Frame?", True)
    NrkSdk.SetBoolArg("Recall Visible Layer?", True)
    NrkSdk.SetIntegerArg("Callout Leader Thickness", 2)
    NrkSdk.SetColorArg("Callout Leader Color", 128, 128, 128)
    NrkSdk.SetIntegerArg("Callout Border Thickness", 2)
    NrkSdk.SetColorArg("Callout Border Color", 0, 0, 255)
    NrkSdk.SetBoolArg("Divide Text with Lines?", False)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_callout_view(collection: str, callout_name: str) -> None:
    """p410"""
    func_name = "Delete Callout View"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Callout View", collection, callout_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def make_a_system_string(str_option: str) -> str:
    """p427"""
    func_name = "Make a System String"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    # Available options:
    # "SA Version", "XIT Filename", "MP Filename", "MP Filename (Full Path)", "Date & Time",
    # "Date", "Date (Short)", "Time", "Key Serial Number", "Company Name",
    # "User Name",
    NrkSdk.SetSystemStringArg("String Content", str_option)
    NrkSdk.SetStringArg("Format String (Optional)", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetStringArg("Resultant String", "")
    if not sValue[0]:
        return ""
    return sValue[1]


def make_a_point_name_runtime_select(user_prompt: str) -> NamedPoint:
    """p435"""
    func_name = "Make a Point Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", user_prompt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    point = NrkSdk.GetPointNameArg("Resultant Point Name", "", "", "")
    log.debug(f"Point: {point}")
    if not point[0]:
        raise ValueError("User selection isn't correct!")

    collection = point[1]
    group = point[2]
    name = point[3]
    return NamedPoint([collection, group, name])


def make_a_point_name_ref_list_from_a_group(collection: str, group: str) -> list[NamedPoint]:
    """p439"""
    func_name = "Make a Point Name Ref List From a Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Group Name", collection, group)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userPtList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if not ptList[0]:
        return []

    points = []
    for i in range(ptList[1].GetLength(0)):
        points.append(NamedPoint(ptList[1][i].split("::")))
    return points


def make_a_point_name_ref_list_runtime_select(user_prompt: str) -> list[NamedPoint]:
    """p440"""
    func_name = "Make a Point Name Ref List - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", user_prompt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userPtList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if not ptList[0]:
        return []

    points = []
    for i in range(ptList[1].GetLength(0)):
        points.append(NamedPoint(ptList[1][i].split("::")))
    return points


def make_a_collection_name_runtime_select(user_prompt: str) -> str:
    """p446"""
    func_name = "Make a Collection Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", user_prompt)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    sValue = NrkSdk.GetCollectionNameArg("Resultant Collection Name", "")
    if not sValue[0]:
        return ""
    return sValue[1]


def make_a_collection_object_name_runtime_select(user_prompt: str, obj_type: str) -> tuple[str, ...]:
    """p450"""
    func_name = "Make a Collection Object Name - Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", user_prompt)

    # Available options:
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud",
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse",
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter",
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh",
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group",
    NrkSdk.SetObjectTypeArg("Object Type", obj_type)

    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return ()
    result = NrkSdk.GetCollectionObjectNameArg("Resultant Collection Object Name", "", "")
    if not result[0]:
        return ()
    return (result[1], result[2])


def make_a_collection_object_name_ref_list_by_type(collection: str, objtype: str) -> list[list[str]]:
    """p454"""
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
    userObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Collection Object Name List", userObjectList)
    if not objectList[0]:
        return []
    objects = []
    for i in range(objectList[1].GetLength(0)):
        objects.append(objectList[1][i].split("::"))  # splits the string in 'collection' and 'object_name'
    return objects


def make_a_relationship_reference_list_wildCard_selection(collection: str, name_relationship: str) -> list[list[str]]:
    """p464"""
    func_name = "Make a Relationship Reference List- WildCard Selection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Collection Wildcard Criteria", collection)
    NrkSdk.SetStringArg("Relationship Wildcard Criteria", name_relationship)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        log.error(f"An empty results was returned for: {collection}::{name_relationship}")
        return []

    userObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Relationship Reference List", userObjectList)
    if not objectList[0]:
        return []
    objects = []
    for i in range(objectList[1].GetLength(0)):
        objects.append(objectList[1][i].split("::"))  # splits the string in 'collection' and 'relationship_name'
    return objects


def make_a_relationship_reference_list_runtime_selection(question: str) -> list[list[str]]:
    """p465"""
    func_name = "Make a Relationship Reference List- Runtime Select"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("User Prompt", question)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    userObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Relationship Reference List", userObjectList)
    if not objectList[0]:
        return []

    objects = []
    for i in range(objectList[1].GetLength(0)):
        objects.append(objectList[1][i].split("::"))  # splits the string in 'collection' and 'relationship_name'
    return objects


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def get_number_of_collections() -> int:
    """p503"""
    func_name = "Get Number of Collections"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    n = NrkSdk.GetIntegerArg("Total Count", 0)
    log.debug(f"n: {n}")
    if not n[0]:
        return 0
    return n[1]


def get_ith_collection_name(i: int) -> str:
    """p504"""
    func_name = "Get i-th Collection Name"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetIntegerArg("Collection Index", i)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    collection = NrkSdk.GetCollectionNameArg("Resultant Name", "")
    log.debug(f"Collection: {collection}")
    if not collection[0]:
        return ""
    return collection[1]


def get_vector_group_properties(collection: str, name_vectorgroup: str) -> dict:
    """p519"""
    func_name = "Get Vector Group Properties"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Vector Group Name", collection, name_vectorgroup)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return {}

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


def set_vector_group_colorization_options_selected(collection: str, name_vectorgroup: str, **kwargs) -> None:
    """p523"""
    func_name = "Set Vector Group Colorization Options (Selected)"
    log.debug(func_name)
    NrkSdk.SetStep("Set Vector Group Colorization Options (Selected)")
    vgNameList = python_list_to_csharp_list([f"{collection}::{name_vectorgroup}"])
    vVectorGroupNameList = sa_py_tools.GetListWrapper(vgNameList)
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


def get_point_coordinate(collection: str, group: str, name: str) -> Point3D:
    """p527"""
    func_name = "Get Point Coordinate"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return Point3D(0.0, 0.0, 0.0)

    Vector = NrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    log.debug(f"Vector: {Vector}")
    if not Vector[0]:
        return Point3D(0.0, 0.0, 0.0)

    return Point3D(Vector[1], Vector[2], Vector[3])


def get_point_to_point_distance(
    collection_p1: str, group_p1: str, name_p1: str, collection_p2: str, group_p2: str, name_p2: str
) -> tuple[Point3D, float]:
    """p533"""
    func_name = "Get Point To Point Distance"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("First Point", collection_p1, group_p1, name_p1)
    NrkSdk.SetPointNameArg("Second Point", collection_p2, group_p2, name_p2)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    # xval = NrkSdk.GetDoubleArg("X Value", 0.0)
    # yval = NrkSdk.GetDoubleArg("Y Value", 0.0)
    # zval = NrkSdk.GetDoubleArg("Z Value", 0.0)
    # mag = NrkSdk.GetDoubleArg("Magnitude", 0.0)
    # returndict = {
    #     "xval": xval[1],
    #     "yval": yval[1],
    #     "zval": zval[1],
    #     "mag": mag[1],
    # }
    # return returndict

    Vector = NrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    mag = NrkSdk.GetDoubleArg("Magnitude", 0.0)
    log.debug(f"Vector: {Vector}, Magnitude: {mag}")
    if not Vector[0]:
        return (Point3D(0.0, 0.0, 0.0), 0.0)

    return (Point3D(Vector[1], Vector[2], Vector[3]), mag)


def set_default_colorization_options() -> None:
    """p569"""
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


def set_vector_group_display_attributes(magnification: float, blotch_size: float, tolerance: float) -> None:
    """p570"""
    func_name = "Set Vector Group Display Attributes"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetBoolArg("Draw Arrowheads?", False)
    NrkSdk.SetBoolArg("Indicate Values?", False)
    NrkSdk.SetDoubleArg("Vector Magnification", magnification)
    NrkSdk.SetIntegerArg("Vector Width", 1)
    NrkSdk.SetBoolArg("Draw Color Blotches?", False)
    NrkSdk.SetDoubleArg("Blotch Size", blotch_size)
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


def transform_object_by_delta_world_transform_operator(objects: list[str], transform: list[float]) -> None:
    """p582"""
    func_name = "Transform Objects by Delta (World Transform Operator)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    # objects is a list of 'collection'::'name_obj' pairs
    objectsList = []
    for item in objects:
        objectsList.append(f"{item[0]}::{item[1]}::Point Group")
    objNameList = python_list_to_csharp_list(objectsList)
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects to Transform", vObjectList)

    # transform is a list of 6 floats: x,y,z,rx,ry,rz
    T = python_list_to_csharp_2D_array(transform)
    scale = 1.0
    vMatrixobj = sa_py_tools.GetListWrapper(T)
    NrkSdk.SetWorldTransformArg("Delta Transform", vMatrixobj, scale)

    NrkSdk.ExecuteStep()
    getResult(func_name)


def transform_objects_by_delta_about_working_frame(objects: list[tuple[str, str]], transform: list[float]) -> None:
    """p583"""
    func_name = "Transform Objects by Delta (About Working Frame)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    # objects is a list of 'collection'::'name_obj' pairs
    objectsList = []
    for item in objects:
        objectsList.append(f"{item[0][0]}::{item[0][1]}::Point Group")
    objectList = python_list_to_csharp_list(objectsList)
    vObjectList = sa_py_tools.GetListWrapper(objectList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects to Transform", vObjectList)

    # transform is a list of 6 floats: x,y,z,rx,ry,rz
    T = python_list_to_csharp_2D_array(transform)
    vMatrixobj = sa_py_tools.GetListWrapper(T)
    NrkSdk.SetTransformArg("Delta Transform", vMatrixobj)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def fit_geometry_to_point_group(
    geomType: str,
    collection_data: str,
    group_data: str,
    collection_result: str,
    name_result: str,
    name_profile: str,
    report_div: bool,
    fit_tol: float,
    out_tol: float,
) -> None:
    """p585"""
    func_name = "Fit Geometry to Point Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)

    # Available options:
    # "Line", "Plane", "Circle", "Sphere", "Cylinder",
    # "Cone", "Paraboloid", "Ellipse", "Slot",
    NrkSdk.SetGeometryTypeArg("Geometry Type", geomType)

    NrkSdk.SetCollectionObjectNameArg("Group To Fit", collection_data, group_data)
    NrkSdk.SetCollectionObjectNameArg("Resulting Object Name", collection_result, name_result)
    NrkSdk.SetStringArg("Fit Profile Name", name_profile)
    NrkSdk.SetBoolArg("Report Deviations", report_div)
    NrkSdk.SetDoubleArg("Fit Interface Tolerance (-1.0 use profile)", fit_tol)
    NrkSdk.SetBoolArg("Ignore Out of Tolerance Points", out_tol)
    NrkSdk.SetCollectionObjectNameArg("Starting Condition Geometry (optional)", "", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def best_fit_transformation_group_to_group(
    collection_ref: str,
    group_ref: str,
    collection_corr: str,
    group_corr: str,
    show_dialog: bool,
    rms_tol: float,
    max_tol: float,
    allow_scale: bool,
) -> dict:
    """p589"""
    func_name = "Best Fit Transformation - Group to Group"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Reference Group", collection_ref, group_ref)
    NrkSdk.SetCollectionObjectNameArg("Corresponding Group", collection_corr, group_corr)
    NrkSdk.SetBoolArg("Show Interface", show_dialog)
    NrkSdk.SetDoubleArg("RMS Tolerance (0.0 for none)", rms_tol)
    NrkSdk.SetDoubleArg("Maximum Absolute Tolerance (0.0 for none)", max_tol)
    NrkSdk.SetBoolArg("Allow Scale", allow_scale)
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

    T = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    trans_in_work = NrkSdk.GetTransformArg("Transform in Working", T)
    results["trans_in_work"] = csharp_array_to_python_2D_list(trans_in_work[1])

    T = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    trans_optimum = NrkSdk.GetWorldTransformArg("Optimum Transform", T, 0.0)
    results["trans_in_world"] = csharp_array_to_python_2D_list(trans_optimum[1])
    results["scale"] = trans_optimum[2]

    results["rms"] = NrkSdk.GetDoubleArg("RMS Deviation", 0.0)[1]
    results["max_abs_dev"] = NrkSdk.GetDoubleArg("Maximum Absolute Deviation", 0.0)[1]
    results["n_unknown"] = NrkSdk.GetIntegerArg("Number of Unknowns", 0.0)[1]
    results["n_equations"] = NrkSdk.GetIntegerArg("Numnber of Equations", 0.0)[1]
    results["robustness"] = NrkSdk.GetDoubleArg("Robustness", 0.0)[1]

    return results


def get_measurement_weather_data(collection: str, group: str, name: str) -> dict:
    """p595"""
    func_name = "Get Measurement Weather Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
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


def get_measurement_auxiliary_data(collection: str, group: str, name: str, name_aux: str) -> dict:
    """p596"""
    func_name = "Get Measurement Auxiliary Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.SetStringArg("Auxiliary Name", name_aux)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    value = NrkSdk.GetDoubleArg("Value", 0.0)
    units = NrkSdk.GetStringArg("Units", "")
    returndict = {
        "value": value[1],
        "units": units[1],
    }
    return returndict


def get_measurement_info_data(collection: str, group: str, name: str) -> list:
    """p599"""
    func_name = "Get Measurement Info Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    value = NrkSdk.GetStringArg("Info Data", "")
    if not value[1]:
        return []

    # Clean the results before returning them
    results = []
    values = value[1].split(";")
    for val in values:
        results.append(val.strip())

    return results


def make_point_to_point_relationship(
    collection_relationship: str,
    name_relationship: str,
    collection_p1: str,
    group_p1: str,
    name_p1: str,
    collection_p2: str,
    group_p2: str,
    name_p2: str,
) -> None:
    """p673"""
    func_name = "Make Point to Point Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetPointNameArg("First Point Name", collection_p1, group_p1, name_p1)
    NrkSdk.SetPointNameArg("Second Point Name", collection_p2, group_p2, name_p2)
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
    collection_relationship: str,
    name_relationship: str,
    collection_nominal: str,
    group_nominal: str,
    collection_measured: str,
    group_measured: str,
) -> None:
    """p686"""
    func_name = "Make Group to Nominal Group Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetCollectionObjectNameArg("Nominal Group Name", collection_nominal, group_nominal)
    NrkSdk.SetCollectionObjectNameArg("Measured Group Name", collection_measured, group_measured)
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
    collection_relationship: str,
    name_relationship: str,
    collection_nominal: str,
    name_nominal: str,
    data: list,
    collection_result: str,
    name_result: str,
) -> None:
    """p689"""
    func_name = "Make Geometry Fit and Compare to Nominal Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetCollectionObjectNameArg("Nominal Geometry", collection_nominal, name_nominal)

    # 'data' is a list of 'collection'::'pointgroupnames'
    objectsList = []
    for item in data:
        objectsList.append(f"{item[0]}::{item[1]}::Point Group")
    objNameList = python_list_to_csharp_list(objectsList)
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Point Groups to Fit", vObjectList)
    NrkSdk.SetCollectionObjectNameArg("Resulting Object Name (Optional)", collection_result, name_result)
    NrkSdk.SetStringArg("Fit Profile Name (Optional)", "")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_relationship(collection: str, name_relationship: str) -> None:
    """p701"""
    func_name = "Delete Relationship"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, name_relationship)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_general_relationship_statistics(collection: str, name_relationship: str) -> dict:
    """p704"""
    func_name = "Get General Relationship Statistics"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, name_relationship)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    results = {}
    results["max_dev"] = NrkSdk.GetDoubleArg("Max Deviation", 0.0)[1]
    results["rms"] = NrkSdk.GetDoubleArg("RMS", 0.0)[1]
    results["has_sign_dev"] = NrkSdk.GetBoolArg("Has Signed Deviation?", False)[1]
    results["sign_max_dev"] = NrkSdk.GetDoubleArg("Signed Max Deviation", 0.0)[1]
    results["sign_min_dev"] = NrkSdk.GetDoubleArg("Signed Min Deviation", 0.0)[1]
    return results


def get_geom_relationship_criteria(collection_relationship: str, name_relationship: str, criteria: str) -> dict:
    """p726"""
    func_name = "Get Geom Relationship Criteria"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)

    # The 'criteria' are explained in the 'MP Command Reference.pdf', pages: 726-727
    NrkSdk.SetStringArg("Criteria", criteria)

    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return {}
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


def set_relationship_associated_data(
    collection_relationship: str, name_relationship: str, method: str, **kwargs
) -> None:
    """p708"""
    func_name = "Set Relationship Associated Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)

    # The function switches it's method based on the variable 'method'.
    # At this point it only excepts 'point_groups' as an input.
    # Individual points, point clouds and objects aren't support for now.
    if method not in ["point_group", "point_groups"]:
        log.error(
            f"This method is only valid with 'point_group' or 'point_groups' as input for now. Input was: {method}"
        )
        raise ValueError(
            f"This method is only valid with 'point_group' or 'point_groups' as input for now. Input was: {method}"
        )

    if method == "points":
        # individual points
        vPointObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
        NrkSdk.SetPointNameRefListArg("Individual Points", vPointObjectList)
    elif method == "point_group":
        if "point_groups_data" not in kwargs:
            log.error("Missing data at point_group!")
            raise ValueError("Missing data at point_group!")

        data = kwargs["point_groups_data"]
        objNameList = python_list_to_csharp_list(
            [f'{data["collection_measured"]}::{data["group_measured"]}::Point Group']
        )
        vObjectList = sa_py_tools.GetListWrapper(objNameList)
        NrkSdk.SetCollectionObjectNameRefListArg("Point Groups", vObjectList)
    elif method == "point_groups":
        # point groups
        if "point_groups_data" not in kwargs:
            log.error("Missing data at point_groups!")
            raise ValueError("Missing data at point_groups!")

        data = kwargs["point_groups_data"]
        objectsList = [
            f'{data["collection_nominals"]}::{data["group_nominals"]}::Point Group',
            f'{data["collection_measured"]}::{data["group_measured"]}::Point Group',
        ]
        objNameList = python_list_to_csharp_list(objectsList)
        vObjectList = sa_py_tools.GetListWrapper(objNameList)
        NrkSdk.SetCollectionObjectNameRefListArg("Point Groups", vObjectList)
    elif method == "point_cloud":
        # point cloud
        vObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
        NrkSdk.SetCollectionObjectNameRefListArg("Point Clouds", vObjectList)
    elif method == "objects":
        # objects
        vObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
        NrkSdk.SetCollectionObjectNameRefListArg("Objects", vObjectList)

    # additional setting
    NrkSdk.SetBoolArg("Ignore Empty Arguments?", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_relationship_associated_data(collection_relationship: str, name_relationship: str) -> dict:
    """p709"""
    # The function only excepts 'groups' as an input.
    # Individual points, point clouds and objects aren't support for now.

    func_name = "Get Relationship Associated Data"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    results = {"relationship_type": "", "individual_points": [], "point_groups": [], "point_clouds": [], "objects": []}

    # relationship_type
    sValue = NrkSdk.GetStringArg("Relationship Type", "")
    # log.debug(f"sValue: {sValue}")
    if sValue[0]:
        results["relationship_type"] = sValue[1]

    # individual_points
    vPointObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    userPtList = NrkSdk.GetPointNameRefListArg("Individual Points", vPointObjectList)
    # log.debug(f"userPtList: {userPtList}")
    if userPtList[0]:
        for i in range(userPtList[1].GetLength(0)):
            # log.debug(f"userPtList {i}: {userPtList[1][i]}")
            results["individual_points"].append(
                userPtList[1][i].split("::")
            )  # splits the string in 'collection' and 'relationship_name'

    # point_groups
    objNameList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    PointGroups = NrkSdk.GetCollectionObjectNameRefListArg("Point Groups", objNameList)
    # log.debug(f"PointGroups: {PointGroups}")
    if PointGroups[0]:
        for i in range(PointGroups[1].GetLength(0)):
            # log.debug(f"PointGroups {i}: {PointGroups[1][i]}")
            results["point_groups"].append(
                PointGroups[1][i].split("::")
            )  # splits the string in 'collection' and 'relationship_name'

    # point_clouds
    vObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    PointClouds = NrkSdk.GetCollectionObjectNameRefListArg("Point Clouds", vObjectList)
    # log.debug(f"PointClouds: {PointClouds}")
    if PointClouds[0]:
        for i in range(PointClouds[1].GetLength(0)):
            log.debug(f"PointClouds {i}: {PointClouds[1][i]}")
            results["point_clouds"].append(PointClouds[1][i])

    # objects
    vObjectList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Objects", vObjectList)
    # log.debug(f"objectList: {objectList}")
    if objectList[0]:
        for i in range(objectList[1].GetLength(0)):
            log.debug(f"object {i}: {objectList[1][i]}")
            results["objects"].append(objectList[1][i])

    return results


def set_relationship_reporting_frame(
    collection_relationship: str, name_relationship: str, collection_frame: str, name_frame: str
) -> None:
    """p723"""
    func_name = "Set Relationship Reporting Frame"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetCollectionObjectNameArg("Reporting Frame", collection_frame, name_frame)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_geom_relationship_criteria(collection_relationship: str, name_relationship: str, criteria_type: str) -> None:
    """p725"""
    func_name = "Set Geom Relationship Criteria"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)

    # Flatness
    if criteria_type == "Flatness":
        NrkSdk.SetStringArg("Criteria", "Flatness")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.1, True, -0.1)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)

    # Centroid Z
    elif criteria_type == "Centroid Z":
        NrkSdk.SetStringArg("Criteria", "Centroid Z")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)

    # Avg Distance Between
    elif criteria_type == "Avg Dist Between":
        NrkSdk.SetStringArg("Criteria", "Avg Dist Between")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)

    # Length
    elif criteria_type == "Length":
        NrkSdk.SetStringArg("Criteria", "Length")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.05, True, -0.05)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    else:
        log.warning("Incorrect criteria type set!")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_geom_relationship_cardinal_points(
    collection_relationship: str, name_relationship: str, name_group: str
) -> None:
    """p734"""
    func_name = "Set Geom Relationship Cardinal Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetBoolArg("Create Cardinal Pts when Fitting?", True)
    NrkSdk.SetBoolArg("Prefix Cardinal Pts name with Rel name?", True)
    NrkSdk.SetStringArg("Cardinal Pts Group Name", name_group)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_geom_relationship_cardinal_points(collection_relationship: str, name_relationship: str) -> list[NamedPoint]:
    """p735"""
    func_name = "Get Geom Relationship Cardinal Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return []

    ptNameList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Cardinal Point Name List", ptNameList)
    if not objectList[0]:
        return []
    points = []
    for i in range(objectList[1].GetLength(0)):
        points.append(NamedPoint(objectList[1][i].split("::")))
    return points


def set_geom_relationship_auto_vectors_nominal_avn(
    collection_relationship: str, name_relationship: str, create_autovectors: bool
) -> None:
    """p739"""
    func_name = "Set Geom Relationship Auto Vectors Nominal (AVN)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetBoolArg("Create Auto Vectors AVN", create_autovectors)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_auto_vectors_fit_avf(
    collection_relationship: str, name_relationship: str, create_autovectors: bool
) -> None:
    """p740"""
    func_name = "Set Relationship Auto Vectors Fit (AVF)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetBoolArg("Create Auto Vectors AVF", create_autovectors)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_desired_meas_count(collection_relationship: str, name_relationship: str, count: int) -> None:
    """p742"""
    func_name = "Set Relationship Desired Meas Count"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
    NrkSdk.SetIntegerArg("Desired Measurement Count", count)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_relationship_tolerance_vector_type(collection_relationship: str, name_relationship: str) -> None:
    """p746"""
    func_name = "Set Relationship Tolerance (Vector Type)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
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
def set_vector_group_report_options(collection: str, name_vectorgroup: str, **kwargs) -> None:
    """p819"""
    func_name = "Set Vector Group Report Options"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Vector Group", collection, name_vectorgroup)
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


def set_relationship_report_options(collection_relationship: str, name_relationship: str, **kwargs) -> None:
    """p820"""
    func_name = "Set Relationship Report Options"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection_relationship, name_relationship)
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


def notify_user_text_array(txt: str, timeout: int = 0) -> None:
    """p854"""
    func_name = "Notify User Text Array"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    stringList = python_list_to_csharp_list([txt])
    vStringList = sa_py_tools.GetListWrapper(stringList)
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
def get_last_instrument_index() -> int:
    """p920"""
    func_name = "Get Last Instrument Index"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    InstID = NrkSdk.GetIntegerArg("Instrument ID", 0)
    if not InstID[0]:
        return -1
    log.debug(f"InstID: {InstID[1]}")
    return InstID[1]


def point_at_target(
    collection_inst: str, id_inst: int, collection_target: str, group_target: str, name_target: str
) -> None:
    """p927"""
    func_name = "Point At Target"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.SetPointNameArg("Target ID", collection_target, group_target, name_target)
    NrkSdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    NrkSdk.ExecuteStep()
    result = getResult_Bare(func_name)
    if result in [-1, 0, 3, 6]:
        log.error(f"Error code was: {result}")
        raise SystemError(f"Failed pointing at point: {collection_target}::{group_target}::{name_target}")
    if result in [1, 2, 4, 5]:
        return


def measure_single_point_here(
    collection_inst: str, id_inst: int, collection: str, group: str, name: str, measure_immediately: bool = False
) -> bool:
    """p928"""
    func_name = "Measure Single Point Here"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.SetPointNameArg("Target ID", "", collection, group, name)
    NrkSdk.SetBoolArg("Measure Immediately", measure_immediately)
    NrkSdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    return True


def stop_active_measurement_mode(collection_inst: str, id_inst: int) -> bool:
    """p934"""
    func_name = "Stop Active Measurement Mode"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    return True


def add_new_instrument(inst_type: str) -> tuple[str, int]:
    """p939"""
    func_name = "Add New Instrument"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetInstTypeNameArg("Instrument Type", inst_type)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    Col_InstID = NrkSdk.GetColInstIdArg("Instrument Added (result)", "", 0)
    log.debug(f"ColInstID: {Col_InstID}")
    if not Col_InstID[0]:
        return ("", -1)
    return (Col_InstID[1], Col_InstID[2])


def initiate_servo_guide(
    collection_inst: str,
    id_inst: int,
    nomPoints: list[Union[NamedPoint, NamedPoint3D]],
    groupname_suffix: str = "",
    targetname_suffix: str = "",
    tolerance: float = 1.0,
) -> bool:
    """p944"""
    func_name = "Initiate Servo-Guide"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    ptList = []
    for point in nomPoints:
        ptList.append(f"{point.collection}::{point.group}::{point.name}")
    ptNameList = python_list_to_csharp_list(ptList)
    vPointObjectList = sa_py_tools.GetListWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg("Nominal Points", vPointObjectList)
    NrkSdk.SetStringArg("Group name suffix", groupname_suffix)
    NrkSdk.SetStringArg("Target name suffix", targetname_suffix)
    NrkSdk.SetDoubleArg("Tolerance", tolerance)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return False
    return True


def watch_point_to_point(
    collection_inst: str, id_inst: int, ref_point: Union[NamedPoint, NamedPoint3D], measure_mode
) -> None:
    """p945"""
    func_name = "Watch Point To Point"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.SetPointNameArg("Reference Point", ref_point.collection, ref_point.group, ref_point.name)
    NrkSdk.SetCollectionObjectNameArg("3 DOF Watch Window Properties", "", "3D Template")
    NrkSdk.SetStringArg("Measurement Mode", measure_mode)
    NrkSdk.SetBoolArg("Pause MP Until Closed", False)
    NrkSdk.SetIntegerArg("Window Top Left X Position", 0)
    NrkSdk.SetIntegerArg("Window Top Left Y Position", 0)
    NrkSdk.SetIntegerArg("Window Width", 500)
    NrkSdk.SetIntegerArg("Window Height", 500)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def watch_point_to_point_with_view_zooming(
    collection_inst: str, id_inst: int, ref_point: Union[NamedPoint, NamedPoint3D]
) -> None:
    """p950"""
    func_name = "Watch Point To Point With View Zooming"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.SetPointNameArg("Reference Point", ref_point.collection, ref_point.group, ref_point.name)
    NrkSdk.SetBoolArg("Update(TRUE),Close(FALSE)", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def start_instrument_interface(
    collection_inst: str, id_inst: int, initialize: bool = True, simulation: bool = False
) -> None:
    """p952"""
    func_name = "Start Instrument Interface"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.SetBoolArg("Initialize at Startup", initialize)
    NrkSdk.SetStringArg("Device IP Address (optional)", "")
    NrkSdk.SetIntegerArg("Interface Type (0=default)", 0)
    NrkSdk.SetBoolArg("Run in Simulation", simulation)
    NrkSdk.SetBoolArg("Allow Start w/o Init Requirements", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def stop_instrument_interface(collection_inst: str, id_inst: int) -> None:
    """p953"""
    func_name = "Stop Instrument Interface"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def verify_instrument_connection(collection_inst: str, id_inst: int) -> bool:
    """p955"""
    func_name = "Verify Instrument Connection"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.ExecuteStep()
    getResult(func_name)
    bValue = NrkSdk.GetBoolArg("Connected?", False)
    log.debug(f"Tracker Connection info: {bValue[1]}")
    if not bValue[0]:
        return False
    return bValue[1]


def configure_and_measure(
    collection_inst: str,
    id_inst: int,
    collection_target: str,
    group_target: str,
    name_target: str,
    profile_name: str,
    measure_immediately: bool,
    wait_for_completion: bool,
    timeout_in_secs: float,
) -> bool:
    """p956"""
    func_name = "Configure and Measure"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.SetPointNameArg("Target Name", collection_target, group_target, name_target)
    NrkSdk.SetStringArg("Measurement Mode", profile_name)
    NrkSdk.SetBoolArg("Measure Immediately", measure_immediately)
    NrkSdk.SetBoolArg("Wait for Completion", wait_for_completion)
    NrkSdk.SetDoubleArg("Timeout in Seconds", timeout_in_secs)
    NrkSdk.ExecuteStep()
    return getResult(func_name)


def measure(collection_inst: str, id_inst: int) -> bool:
    """p958"""
    func_name = "Measure"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.ExecuteStep()
    return getResult(func_name)


def compute_CTE_scale_factor(cte: float, parttemp: float) -> float:
    """p979"""
    func_name = "Compute CTE Scale Factor"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    NrkSdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    NrkSdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    scaleFactor = NrkSdk.GetDoubleArg("Scale Factor", 0.0)
    if not scaleFactor[0]:
        return 0.0
    return scaleFactor[1]


def set_instrument_scale_absolute(collection_inst: str, id_inst: int, scale_factor: float) -> None:
    """p981"""
    func_name = "Set (absolute) Instrument Scale Factor (CAUTION!)"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.SetDoubleArg("Scale Factor", scale_factor)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def move_measurement_observation(
    collection: str, group: str, name: str, index: int, collection_dest: str, group_dest: str, name_dest: str
) -> None:
    "p943"
    func_name = "Move Measurement Observation"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Source Point Name", collection, group, name)
    NrkSdk.SetIntegerArg("Observation index", index)
    NrkSdk.SetBoolArg("Delete point if no measurements remain?", False)
    NrkSdk.SetPointNameArg("Destination Point Name", collection_dest, group_dest, name_dest)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def instrument_operational_check(collection_inst: str, id_inst: int, check_type: str) -> bool:
    """p986"""
    func_name = "Instrument Operational Check"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument to Check", collection_inst, id_inst)

    # For the available check types see documention "MP Command Reference.pdf"
    NrkSdk.SetStringArg("Check Type", check_type)

    NrkSdk.ExecuteStep()
    return getResult(func_name)


def get_number_of_observations_on_target(collection: str, group: str, name: str) -> int:
    """p998"""
    func_name = "Get Number of Observations on Target"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return 0

    value = NrkSdk.GetIntegerArg("Number of Shots", 0)
    if not value[0]:
        return 0
    return value[1]


def get_targets_measured_by_instrument(collection: str, instrument_id: int) -> list[NamedPoint]:
    """p1000"""
    func_name = "Get Targets Measured by Instrument"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Measuring Instrument ID", collection, instrument_id)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        log.error("Executing the function resluted with an error.")
        return []

    userPtList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    ptList = NrkSdk.GetPointNameRefListArg("Points Measured by Instrument", userPtList)
    if not ptList[0]:
        log.warning("Tracker doesn't have measured points.")
        return []

    points = []
    for i in range(ptList[1].GetLength(0)):
        points.append(NamedPoint(ptList[1][i].split("::")))

    return points


def get_observation_info(collection: str, group: str, name: str, index=0) -> dict:
    """p1002"""
    func_name = "Get Observation Info"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.SetIntegerArg("Observation Index", index)
    NrkSdk.ExecuteStep()
    result = getResult_Bare(func_name)

    if not result[0]:
        log.debug("No result.")
        return {}

    if result[1] != 2:
        log.debug(f"Result != 2: {result}")
        return {}

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


def set_instrument_measurement_mode_profile(collection_inst: str, id_inst: int, mode_profile: str) -> None:
    """p1005"""
    func_name = "Set Instrument Measurement Mode/Profile"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument to set", collection_inst, id_inst)
    NrkSdk.SetStringArg("Mode/Profile", mode_profile)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_instrument_target_status(collection_inst: str, id_inst: int) -> dict:
    """p1013"""
    func_name = "Get Instrument Target Status"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection_inst, id_inst)
    NrkSdk.ExecuteStep()
    if not getResult(func_name):
        return {}

    isLocked = NrkSdk.GetBoolArg("Is Locked?", False)[1]
    activeTarget = NrkSdk.GetStringArg("Name", "")[1]
    nFaces = NrkSdk.GetIntegerArg("Number of Faces", 0)[1]
    lockedFace = NrkSdk.GetIntegerArg("Locked Face", 0)[1]
    results = {
        "isLocked": isLocked,
        "activeTarget": activeTarget,
        "nFaces": nFaces,
        "lockedFace": lockedFace,
    }
    return results


def auto_measure_points(
    collection_inst: str,
    id_inst: int,
    collection_ref: str,
    group_ref: str,
    collection_measured: str,
    group_measured: str,
) -> None:
    """p1017"""
    func_name = "Auto Measure Points"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.SetCollectionObjectNameArg("Reference Group Name", collection_ref, group_ref)
    NrkSdk.SetCollectionObjectNameArg("Actuals Group Name (to be measured)", collection_measured, group_measured)
    NrkSdk.SetBoolArg("Force use of existing group?", False)
    NrkSdk.SetBoolArg("Show complete dialog?", True)
    NrkSdk.SetBoolArg("Wait for Completion?", True)
    NrkSdk.SetBoolArg("Auto Start?", False)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def auto_correspond_closest_point(
    collection_inst: str,
    id_inst: int,
    collection_ref: str,
    group_ref: str,
    collection_measured: str,
    group_measured: str,
) -> None:
    """p1021"""
    func_name = "Auto-Correspond Closest Point"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.SetCollectionObjectNameArg("Reference Group Name", collection_ref, group_ref)
    NrkSdk.SetCollectionObjectNameArg("Actuals Group Name (to be measured)", collection_measured, group_measured)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def auto_correspond_with_proximity_trigger(
    collection_inst: str,
    id_inst: int,
    collection_nom: str,
    group_nom: str,
    collection_measured: str,
    group_measured: str,
    proximity: float,
) -> None:
    """p1022"""
    func_name = "Auto-Correspond with Proximity Trigger"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetColInstIdArg("Instrument ID", collection_inst, id_inst)
    NrkSdk.SetCollectionObjectNameArg("Nominal Point Group or Vector Group", collection_nom, group_nom)
    NrkSdk.SetCollectionObjectNameArg("Results Point Group for measurements", collection_measured, group_measured)
    NrkSdk.SetDoubleArg("Point distance threshold", proximity)
    NrkSdk.SetDoubleArg("Vector axis threshold", 0.250000)
    NrkSdk.SetBoolArg("Project results to nominal vector", False)
    NrkSdk.SetDoubleArg("Warbler ramp start zone distance", 12.000000)
    NrkSdk.SetBoolArg("Show Watch window on startup", False)
    NrkSdk.SetVectorGroupNameArg("Vector Group to make while Measuring (blank means ignore)", "")
    NrkSdk.SetBoolArg("Make unmeasured group when done", False)
    NrkSdk.SetBoolArg("Measure each point only once", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


# ################################
# Chapter 13 - Robot Operations ##
# ################################


# ##################################
# Chapter 14 - Utility Operations ##
# ##################################
def delete_folder(foldername: str) -> None:
    """p1116"""
    func_name = "Delete Folder"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Folder Path", foldername)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def move_collection_to_folder(collection: str, folder: str) -> None:
    """p1117"""
    func_name = "Move Collection to Folder"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def get_folders_by_wildcard(search: str) -> list[str]:
    """p1119"""
    func_name = "Get Folders by Wildcard"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Search String", search)
    NrkSdk.SetBoolArg("Case Sensitive Search", False)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    stringList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    vStringList = NrkSdk.GetStringRefListArg("Folder List", stringList)
    if vStringList[0]:
        return []
    folders = []
    for i in range(vStringList[1].GetLength(0)):
        folders.append(vStringList[1][i])
    return folders


def get_folder_collections(folder: str) -> list[str]:
    """p1122"""
    func_name = "Get Folder Collections"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(func_name)

    stringList = sa_py_tools.GetListWrapper(python_list_to_csharp_list([]))
    vStringList = NrkSdk.GetStringRefListArg("Collection List", stringList)
    if not vStringList[0]:
        return []
    folders = []
    for i in range(vStringList[1].GetLength(0)):
        folders.append(vStringList[1][i])
    return folders


def set_collection_notes(collection: str, notes: str) -> None:
    """p1125"""
    func_name = "Set Collection Notes"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    stringList = python_list_to_csharp_list([notes])
    vStringList = sa_py_tools.GetListWrapper(stringList)
    NrkSdk.SetEditTextArg("Notes", vStringList)
    NrkSdk.SetBoolArg("Append? (FALSE = Overwrite)", True)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_working_frame(collection: str, name: str) -> None:
    """p1142"""
    func_name = "Set Working Frame"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameArg("New Working Frame Name", collection, name)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_objects(collection: str, name: str, objtype: str) -> None:
    """p1151"""
    func_name = "Delete Objects"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    objectsList = [
        f"{collection}::{name}::{objtype}",
    ]
    objNameList = python_list_to_csharp_list(objectsList)
    vObjectList = sa_py_tools.GetListWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Object Names", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(func_name)


def delete_items():
    """p1152"""
    func_name = "Delete Items"
    log.debug(func_name)
    NrkSdk.SetStep(func_name)
    NrkSdk.SetCollectionObjectNameRefListArg("Item List")
    NrkSdk.ExecuteStep()
    getResult(func_name)


def set_interaction_mode(sa_interaction_mode: str, mp_interaction_mode: str, mp_dialog_interaction_mode: str) -> None:
    """p1180"""
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
