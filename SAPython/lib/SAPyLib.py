"""
SAPyLib = Spatial Analyzer Python .NET library

This library provides the communication layer to the SA SDK .NET dll.
Author: L. Ververgaard
Date: 2020-04-08
Version: 1
"""
import sys
import os
import clr

basepath = r"C:\Analyzer Data\Scripts\SAPython"

# Set the debugging reporting value
DEBUG = True

clr.AddReference("System")
clr.AddReference("System.Reflection")
import System
import System.Reflection
from System.Collections.Generic import List


# SA SDK dll
sa_sdk_dll_file = os.path.join(basepath, r"dll\Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
sdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = sdk.Connect("127.0.0.1")
if not SAConnected:
    raise IOError('Connection to SA SDK failed!')
    sys.exit(1)


# ###############
# Base methods ##
# ###############
def dprint(val):
    """Print a debugging value."""
    if DEBUG:
        print(f'\tDEBUG: {val}')

def fprint(val):
    """Print the function name."""
    if DEBUG:
        print(f'\nFUNCTION: {val}')

def getResult(fname):
    """Get the methods execution result."""
    # result = getIntRef()
    boolean, result = sdk.GetMPStepResult(0)
    dprint(f'{fname}: {boolean}, {result}')
    if result == -1:
        # SDKERROR = -1
        raise SystemError('Execution raised: SDKERROR!')
    elif result == 0:
        # UNDONE = 0
        dprint('Execution: undone.')
        raise IOError()
    elif result == 1:
        # INPROGRESS = 1
        dprint('Execution: inprogress.')
    elif result == 2:
        # DONESUCCESS = 2
        return True
    elif result == 3:
        # DONEFATALERROR = 3
        dprint('Execution: FAILED!')
        raise IOError()
    elif result == 4:
        # DONEMINORERROR = 4
        dprint('Execution: FAILED - minor error!')
        return True
    elif result == 5:
        # CURRENTTASK = 5
        dprint('Execution: current task')
        raise IOError()
    elif result == 6:
        # UNKNOWN = 6
        dprint('I have no clue!')
        raise IOError()


class Point3D:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z

class NamedPoint3D:
    def __init__(self, n, x, y, z):
        self.name = n
        self.X = x
        self.Y = y
        self.Z = z

class CTE:
    AluminumCTE_1DegF = 0.0000131
    SteelCTE_1DegF = 0.0000065
    CarbonFiberCTE_1DegF = 0


# ##############################
# Chapter 2 - File Operations ##
# ##############################


# ######################################
# Chapter 3 - Process Flow Operations ##
# ######################################
def object_existence_test(ColObjName):
    """p111"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

    sdk.ExecuteStep()
    getResult(fname)


def ask_for_string(question, *args, **kwargs):
    """p117"""

    if 'initialanswer' in kwargs:
        initialanswer = kwargs['initialanswer']
    else:
        initialanswer = ''
    
    fname = 'Ask for String'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetStringArg("Question to ask", question)
    sdk.SetBoolArg("Password Entry?", False)
    sdk.SetStringArg("Initial Answer", initialanswer)
    sdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    sdk.ExecuteStep()
    getResult(fname)
    answer = sdk.GetStringArg("Answer", '')
    return answer[1]


def ask_for_string_pulldown(question, answers):
    """p118"""
    fname = 'Ask for String (Pull-Down Version)'
    fprint(fname)
    sdk.SetStep(fname)
    # question section
    questionList = [question, ]
    QvStringList = System.Runtime.InteropServices.VariantWrapper(questionList)
    sdk.SetStringRefListArg("Question or Statement", QvStringList)
    # anwsers section
    AvStringList = System.Runtime.InteropServices.VariantWrapper(answers)
    sdk.SetStringRefListArg("Possible Answers", AvStringList)
    sdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    sdk.ExecuteStep()
    getResult(fname)
    answer = sdk.GetStringArg("Answer", '')
    return answer[1]


def ask_for_user_decision_extended(question, choices):
    """p122"""
    fname = 'Ask for User Decision Extended'
    fprint(fname)
    sdk.SetStep(fname)
    stringList = [question, ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    sdk.SetEditTextArg("Question or Statement", vStringList)
    sdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    #sdk.NOT_SUPPORTED("Button Answers", choices)
    sdk.ExecuteStep()
    getResult(fname)
    selected_answer = 0
    return selected_answer


# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################


# ###########################
# Chapter 5 - View Control ##
# ###########################
def hide_objects(collection, objects, name):
    """p153"""
    fname = 'Hide Objects'
    fprint(fname)
    sdk.SetStep(fname)
    objNameList = [f'{collection}::{objects}::{name}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Objects To Hide", vObjectList)
    sdk.ExecuteStep()
    getResult(fname)


def show_hide_objects(ColName, ObjType, hide):
    """p154"""
    # Hide = hide/True
    # Show = hide/False
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

    sdk.ExecuteStep()
    getResult(fname)


def show_objects(collection, objects, name):
    """p155"""
    fname = 'Show Objects'
    fprint(fname)
    sdk.SetStep(fname)
    objNameList = [f'{collection}::{objects}::{name}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Objects To Show", vObjectList)
    sdk.ExecuteStep()
    getResult(fname)


def show_hide_callout_view(CalloutName, show):
    """p159"""
    # Show = show/True
    # Hide = show/False
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

    sdk.ExecuteStep()
    getResult(fname)


def hide_all_callout_views():
    """p160"""
    fname = 'Hide All Callout Views'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.ExecuteStep()
    getResult(fname)


def center_graphics_about_objects(ColWild, ObjWild):
    """p170"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

    sdk.ExecuteStep()
    getResult(fname)


# ######################################
# Chapter 6 - Cloud Viewer Operations ##
# ######################################


# ######################################
# Chapter 7 - Construction Operations ##
# ######################################
def rename_point(orgCol, orgGrp, orgName, newCol, newGrp, newName):
    """p192"""
    fname = 'Rename Point'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetPointNameArg("Original Point Name", orgCol, orgGrp, orgName)
    sdk.SetPointNameArg("New Point Name", newCol, newGrp, newName)
    sdk.SetBoolArg("Overwrite if exists?", False)
    sdk.ExecuteStep()
    getResult(fname)


def rename_collection(fromName, toName):
    """p194"""
    fname = "Rename Collection"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Original Collection Name", fromName)
    sdk.SetCollectionNameArg("New Collection Name", toName)
    sdk.ExecuteStep()
    getResult(fname)


def delete_points(collection, group, point):
    """p197"""
    fname = 'Delete Points'
    fprint(fname)
    sdk.SetStep(fname)
    ptNameList = [f'{collection}::{group}::{point}', ]
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    sdk.SetPointNameRefListArg('Point Names', vPointObjectList)
    sdk.ExecuteStep()
    getResult(fname)


def delete_points_wildcard_selection(Col, Group, ObjType, pName):
    """p198"""
    fname = "Delete Points WildCard Selection"
    fprint(fname)
    sdk.SetStep(fname)
    objNameList = [f'{Col}::{Group}::{ObjType}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Groups to Delete From", vObjectList)
    sdk.SetPointNameArg("WildCard Selection Names", "*", "*", pName)
    sdk.ExecuteStep()
    getResult(fname)


def construct_objects_from_surface_faces_runtime_select(*args, **kwargs):
    """p199"""
    fname = 'Construct Objects From Surface Faces - Runtime Select'
    fprint(fname)
    sdk.SetStep(fname)
    if 'facetype' in kwargs:
        facetype = kwargs['facetype']
        if not facetype:
            print('No facetype selected!')
            sys.exit(1)
    # Set correct states
    state = [False, False, False, False, False, False, False]
    if facetype == 'plane':
        state[0] = True
    elif facetype == 'cylinder':
        state[1] = True
    elif facetype == 'spere':
        state[2] = True
    elif facetype == 'cone':
        state[3] = True
    elif facetype == 'line':
        state[4] = True
    elif facetype == 'point':
        state[5] = True
    elif facetype == 'circle':
        state[6] = True
    # Set the correct type
    sdk.SetBoolArg("Construct Planes?", state[0])
    sdk.SetBoolArg("Construct Cylinders?", state[1])
    sdk.SetBoolArg("Construct Spheres?", state[2])
    sdk.SetBoolArg("Construct Cones?", state[3])
    sdk.SetBoolArg("Construct Lines?", state[4])
    sdk.SetBoolArg("Construct Points?", state[5])
    sdk.SetBoolArg("Construct Circles?", state[6])
    sdk.ExecuteStep()


def set_or_construct_default_collection(colName):
    """p201"""
    fname = "Set (or construct) default collection"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection Name", colName)
    sdk.ExecuteStep()
    getResult(fname)


def construct_collection(name, makeDefault):
    """p202"""
    # default collection = makeDefault/True
    # collection = makeDefault/False
    fname = "Construct Collection"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection Name", name)
    sdk.SetStringArg("Folder Path", "")
    sdk.SetBoolArg("Make Default Collection?", makeDefault)
    sdk.ExecuteStep()
    getResult(fname)


def get_active_collection():
    """p203"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

    sdk.ExecuteStep()
    getResult(fname)
    activeCollection = ''
    return activeCollection


def delete_collection(collection):
    """p204"""
    fname = "Delete Collection"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    sdk.ExecuteStep()
    getResult(fname)


def construct_a_point(collection, group, name, x, y, z):
    """p208"""
    fname = "Construct a Point in Working Coordinates"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetPointNameArg("Point Name", collection, group, name)
    sdk.SetVectorArg("Working Coordinates", x, y, z)
    sdk.ExecuteStep()
    getResult(fname)


def construct_line_2_points(name, fCol, fGrp, fTarg, sCol, sGrp, sTarg):
    """p262"""
    fname = 'Construct Line 2 Points'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Line Name", "", name)
    sdk.SetPointNameArg("First Point", fCol, fGrp, fTarg)
    sdk.SetPointNameArg("Second Point", sCol, sGrp, sTarg)
    sdk.ExecuteStep()
    getResult(fname)


def construct_frame_known_origin_object_direction_object_direction(
                                                        ptCol, ptGrp, ptTarg,
                                                        x, y, z,
                                                        priCol, priObj, priAxisDef,
                                                        secCol, secObj, secAxisDef,
                                                        frameCol, frameName
                                                        ):
    """p327"""
    fname = 'Construct Frame, Known Origin, Object Direction, Object Direction'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetPointNameArg("Known Point", ptCol, ptGrp, ptTarg)
    sdk.SetVectorArg("Known Point Value in New Frame", x, y, z)
    sdk.SetCollectionObjectNameArg("Primary Axis Object", priCol, priObj)
    # Available options: 
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis", 
    # "-Z Axis", 
    sdk.SetAxisNameArg("Primary Axis Defines Which Axis", priAxisDef)
    sdk.SetCollectionObjectNameArg("Secondary Axis Object", secCol, secObj)
    # Available options: 
    # "+X Axis", "-X Axis", "+Y Axis", "-Y Axis", "+Z Axis", 
    # "-Z Axis", 
    sdk.SetAxisNameArg("Secondary Axis Defines Which Axis", secAxisDef)
    sdk.SetCollectionObjectNameArg("Frame Name (Optional)", frameCol, frameName)
    sdk.ExecuteStep()
    getResult(fname)


def create_relationship_callout():
    """p363"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)
    sdk.__temp__()
    sdk.ExecuteStep()
    getResult(fname)


def create_text_callout():
    """p365"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)
    sdk.__temp__()
    sdk.ExecuteStep()
    getResult(fname)


def set_default_callout_view_properties():
    """p372"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)
    sdk.__temp__()
    sdk.ExecuteStep()
    getResult(fname)


def make_a_point_name_runtime_select(txt):
    """p398"""
    fname = 'Make a Point Name - Runtime Select'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetStringArg("User Prompt", txt)
    sdk.ExecuteStep()
    getResult(fname)
    point = sdk.GetPointNameArg("Resultant Point Name", '', '', '')
    print(point)
    if point[0]:
        sCol = point[1]
        sGrp = point[2]
        sTarg = point[3]
        return (sCol, sGrp, sTarg)
    else:
        return False


def make_a_point_name_ref_list_from_a_group(collection, group):
    """p401"""
    fname = 'Make a Point Name Ref List From a Group'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Group Name", collection, group)
    sdk.ExecuteStep()
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = sdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    print(f'ptList: {ptList}')
    if ptList[0]:
        print(f'ptList[1]: {ptList[1]}')
        points = ptList[1]
        print(f'p: {points}')
        newPtList = []
        for i in range(points.GetLength(0)):
            temp = points[i].split('::')
            newPtList.append(temp)
        return newPtList
    else:
        return False


def make_a_point_name_ref_list_runtime_select(prompt):
    """p402"""
    fname = 'Make a Point Name Ref List - Runtime Select'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetStringArg("User Prompt", prompt)
    sdk.ExecuteStep( )
    getResult(fname)
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = sdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    print(f'ptList: {ptList}')
    if ptList[0]:
        print(f'ptList[1]: {ptList[1]}')
        points = ptList[1]
        print(f'p: {points}')
        newPtList = []
        for i in range(points.GetLength(0)):
            temp = points[i].split('::')
            newPtList.append(temp)
        return newPtList
    else:
        return False


def make_a_collection_object_name_ref_list_by_type(collection, objtype):
    """p416"""
    fname = 'Make a Collection Object Name Ref List - By Type'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetStringArg("Collection", collection)
    # Available options: 
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh", 
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group", 
    sdk.SetObjectTypeArg("Object Type", objtype)
    sdk.ExecuteStep()
    getResult(fname)
    userObjectList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = sdk.GetCollectionObjectNameRefListArg("Resultant Collection Object Name List", userObjectList)
    print(f'objectList: {objectList}')
    if objectList[0]:
        print(f'objectList[1]: {objectList[1]}')
        objects = objectList[1]
        objectsList = []
        for i in range(objects.GetLength(0)):
            temp = objects[i].split('::')
            objectsList.append(temp)
        return objectsList
    else:
        return False


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def get_number_of_collections():
    """p465"""
    fname = 'Get Number of Collections'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.ExecuteStep()
    n = sdk.GetIntegerArg('Total Count', 0)
    if not n:
        return False
    else:
        return n


def get_ith_collection_name(i):
    """p466"""
    fname = 'Get i-th Collection Name'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetIntegerArg('Collection Index', i)
    sdk.ExecuteStep()
    collection = sdk.GetCollectionNameArg('Resultant Name', '')
    print(f'Collection: {collection}')
    if not collection:
        return False
    else:
        return collection


def get_point_coordinate(collection, group, pointname):
    """p489"""
    fname = 'Get Point Coordinate'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetPointNameArg("Point Name", collection, group, pointname)
    sdk.ExecuteStep()
    Vector = sdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    print(f'Vector: {Vector}')
    # xVal = sdk.GetDoubleArg("X Value", 0.0)
    # yVal = sdk.GetDoubleArg("Y Value", 0.0)
    # zVal = sdk.GetDoubleArg("Z Value", 0.0)
    if Vector[0]:
        xVal = Vector[1]
        yVal = Vector[2]
        zVal = Vector[3]
        return (xVal, yVal, zVal)
    else:
        return False


# ########################################
# --- NOT DIRECTLY NEEDED --- 
# CAN BE DONE WITH NORMAL PYTHON FOR LOOP
# ########################################
# def get_ith_point_name_from_point_name_ref_list_iterator(ptList, i):
#     """p498"""
#     fname = 'Get i-th Point Name From Point Name Ref List (Iterator)'
#     fprint(fname)
#     sdk.SetStep(fname)
#     vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptList)
#     sdk.SetPointNameRefListArg("Point Name List", vPointObjectList)
#     sdk.SetIntegerArg("Point Name Index", i+1)
#     sdk.ExecuteStep()
#     getResult(fname)
#     sCol = sdk.GetStringArg("Collection", '')
#     sGrp = sdk.GetStringArg("Group", '')
#     sTarg = sdk.GetStringArg("Target", '')
#     pointname = sdk.GetPointNameArg("Resulting Point Name", '', '', '')
#     print(f'sCol: {sCol}')
#     print(f'sGrp: {sGrp}')
#     print(f'sTarg: {sTarg}')
#     print(f'Pointname: {pointname}')
#     return (sCol, sGrp, sTarg, pointname)


def fit_geometry_to_point_group(geomType,
                                dataCol, dataGroup,
                                resultCol, resultName,
                                profilename, reportdiv, fittol, outtol
                                ):
    """p547"""
    fname = 'Fit Geometry to Point Group'
    fprint(fname)
    sdk.SetStep(fname)
    # Available options: 
    #     "Line", "Plane", "Circle", "Sphere", "Cylinder", 
    #     "Cone", "Paraboloid", "Ellipse", "Slot", 
    sdk.SetGeometryTypeArg("Geometry Type", geomType)
    sdk.SetCollectionObjectNameArg("Group To Fit", dataCol, dataGroup)
    sdk.SetCollectionObjectNameArg("Resulting Object Name", resultCol, resultName)
    sdk.SetStringArg("Fit Profile Name", profilename)
    sdk.SetBoolArg("Report Deviations", reportdiv)
    sdk.SetDoubleArg("Fit Interface Tolerance (-1.0 use profile)", fittol)
    sdk.SetBoolArg("Ignore Out of Tolerance Points", outtol)
    sdk.SetCollectionObjectNameArg("Starting Condition Geometry (optional)", "", "")
    sdk.ExecuteStep( )
    getResult(fname)


def best_fit_group_to_group(refCollection, refGroup, 
                            corCollection, corGroup, 
                            showDialog, rmsTol, maxTol, allowScale):
    """p551"""
    fname = "Best Fit Transformation - Group to Group"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Reference Group", refCollection, refGroup)
    sdk.SetCollectionObjectNameArg("Corresponding Group", corCollection, corGroup)
    sdk.SetBoolArg("Show Interface", showDialog)
    sdk.SetDoubleArg("RMS Tolerance (0.0 for none)", rmsTol)
    sdk.SetDoubleArg("Maximum Absolute Tolerance (0.0 for none)", maxTol)
    sdk.SetBoolArg("Allow Scale", allowScale)
    sdk.SetBoolArg("Allow X", True)
    sdk.SetBoolArg("Allow Y", True)
    sdk.SetBoolArg("Allow Z", True)
    sdk.SetBoolArg("Allow Rx", True)
    sdk.SetBoolArg("Allow Ry", True)
    sdk.SetBoolArg("Allow Rz", True)
    sdk.SetBoolArg("Lock Degrees of Freedom", False)
    sdk.SetFilePathArg("File Path for CSV Text Report (requires Show Interface = TRUE)",
                       "", False)
    sdk.ExecuteStep()
    boolean, result = sdk.GetMPStepResult(0)
    dprint(f'{fname}: {boolean}, {result}')
    return boolean


def make_group_to_nominal_group_relationship(relCol, relName,
                                             nomCol, nomGrp,
                                             meaCol, meaGrp,
                                             ):
    """p646"""
    fname = 'Make Group to Nominal Group Relationship'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    sdk.SetCollectionObjectNameArg("Nominal Group Name", nomCol, nomGrp)
    sdk.SetCollectionObjectNameArg("Measured Group Name", meaCol, meaGrp)
    sdk.SetBoolArg("Auto Update a Vector Group?", False)
    sdk.SetBoolArg("Use Closest Point?", True)
    sdk.SetBoolArg("Display Closest Point Watch Window?", False)
    sdk.SetBoolArg("Use View Zooming With Proximity?", False)
    sdk.SetBoolArg("Ignore Points Beyond Threshold?", False)
    sdk.SetDoubleArg("Proximity Threshold?", 0.01)
    sdk.SetToleranceVectorOptionsArg(
        "Tolerance",
        False, 0.0, False, 0.0, False, 0.0, True, 1.5,
        False, 0.0, False, 0.0, False, 0.0, True, 0.0
        )
    sdk.SetToleranceVectorOptionsArg(
        "Constraint",
        True, 0.0, True, 0.0, True, 0.0, False, 0.0,
        True, 0.0, True, 0.0, True, 0.0, False, 0.0
        )
    sdk.SetDoubleArg("Fit Weight", 1.0)
    sdk.ExecuteStep()
    getResult(fname)


def make_geometry_fit_and_compare_to_nominal_relationship(relatCol, relatName,
                                                          nomCol, nomName,
                                                          data,
                                                          resultCol, resultName
                                                          ):
    """p649"""
    fname = 'Make Geometry Fit and Compare to Nominal Relationship'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relatCol, relatName)
    sdk.SetCollectionObjectNameArg("Nominal Geometry", nomCol, nomName)
    objNameList = []
    for item in data:
        objNameList.append(f'{item[0]}::{item[1]}::Point Group')
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Point Groups to Fit", vObjectList)
    sdk.SetCollectionObjectNameArg("Resulting Object Name (Optional)", resultCol, resultName)
    sdk.SetStringArg("Fit Profile Name (Optional)", "")
    sdk.ExecuteStep()
    getResult(fname)


def set_relationship_associated_data(collection, group, collectionmeasured, groupmeasured):
    """p664"""
    # The function only excepts 'groups' as an input.
    # Individual points, point clouds and objects aren't support for now.
    fname = 'Set Relationship Associated Data'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", collection, group)
    # individual points
    ptNameList = []
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    sdk.SetPointNameRefListArg("Individual Points", vPointObjectList)
    # point groups
    objNameList = [f'{collectionmeasured}::{groupmeasured}::Point Group', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Point Groups", vObjectList)
    # point cloud
    objNameList = []
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Point Clouds", vObjectList)
    # objects
    objNameList = []
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Objects", vObjectList)
    # additional setting
    sdk.SetBoolArg("Ignore Empty Arguments?", True)
    sdk.ExecuteStep()
    getResult(fname)


def set_relationship_reporting_frame(relCol, relGrp, frmCol, frmName):
    """p679"""
    fname = 'Set Relationship Reporting Frame'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    sdk.SetCollectionObjectNameArg("Reporting Frame", frmCol, frmName)
    sdk.ExecuteStep()
    getResult(fname)


def set_geom_relationship_criteria(relCol, relName, criteriatype):
    """p680"""
    fname = 'Set Geom Relationship Criteria'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    if criteriatype == 'Flatness':
        sdk.SetStringArg("Criteria", "Flatness")
        sdk.SetBoolArg("Show in Report", True)
        sdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.1, True, 0.1)
        sdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        sdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    elif criteriatype == 'Centroid Z':
        sdk.SetCollectionObjectNameArg("Relationship Name", "", "")
        sdk.SetStringArg("Criteria", "Centroid Z")
        sdk.SetBoolArg("Show in Report", True)
        sdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, 3.0)
        sdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        sdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    elif criteriatype == 'Avg Dist Between':
        sdk.SetCollectionObjectNameArg("Relationship Name", "", "")
        sdk.SetStringArg("Criteria", "Avg Dist Between")
        sdk.SetBoolArg("Show in Report", True)
        sdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, 3.0)
        sdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        sdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    else:
        print('incorrect criteria type set')
    sdk.ExecuteStep()
    getResult(fname)


def set_relationship_desired_meas_count(relCol, relName, count):
    """p693"""
    fname = 'Set Relationship Desired Meas Count'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    sdk.SetIntegerArg("Desired Measurement Count", count)
    sdk.ExecuteStep()
    getResult(fname)


# ###################################
# Chapter 9 - Reporting Operations ##
# ###################################
def set_relationship_report_options(relCol, relGrp):
    """p770"""
    fname = 'Set Relationship Report Options'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    sdk.SetPointDeltaReportOptionsArg("Report Options", "Cartesian", "Single", True, True, True, True, True, True, False, True, True, True)
    sdk.ExecuteStep()
    getResult(fname)


def notify_user_text_array(txt, timeout=0):
    """p804"""
    fname = 'Notify User Text Array'
    fprint(fname)
    sdk.SetStep(fname)
    stringList = [txt, ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    sdk.SetEditTextArg("Notification Text", vStringList)
    sdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    sdk.SetBoolArg("Auto expand to fit text?", False)
    sdk.SetIntegerArg("Display Timeout", timeout)
    sdk.ExecuteStep( )
    getResult(fname)


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
    fname = 'Get Last Instrument Index'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.ExecuteStep()
    getResult(fname)
    # instId = sdk.GetIntegerArg("Instrument ID", 0)
    ColInstID = sdk.GetColInstIdArg("Instrument ID", '', 0)
    # print(f'\tinstId: {instId}')
    print(f'\tColInstID: {ColInstID}')
    return (ColInstID[1], ColInstID[2])


def point_at_target(instCol, instId, collection, group, name):
    """p877"""
    fname = "Point At Target"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument ID", instCol, instId)
    sdk.SetPointNameArg("Target ID", collection, group, name)
    sdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    sdk.ExecuteStep()
    getResult(fname)


def add_new_instrument(instName):
    """p889"""
    fname = "Add New Instrument"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetInstTypeNameArg("Instrument Type", instName)
    sdk.ExecuteStep()
    getResult(fname)
    ColInstID = sdk.GetColInstIdArg("Instrument Added (result)", '', 0)
    print(f'\tColInstID: {ColInstID}')
    return ColInstID[1], ColInstID[2]


def start_instrument(InstCol, InstID, initialize, simulation):
    """p902"""
    fname = "Start Instrument Interface"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", InstCol, InstID)
    sdk.SetBoolArg("Initialize at Startup", initialize)
    sdk.SetStringArg("Device IP Address (optional)", "")
    sdk.SetIntegerArg("Interface Type (0=default)", 0)
    sdk.SetBoolArg("Run in Simulation", simulation)
    sdk.ExecuteStep()
    getResult(fname)


def stop_instrument(collection, instid):
    """p903"""
    fname = "Stop Instrument Interface"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", collection, instid)
    sdk.ExecuteStep()
    getResult(fname)


def configure_and_measure(instrumentCollection, instId, 
                          targetCollection, group, name, 
                          profileName, measureImmediately, 
                          waitForCompletion, timeoutInSecs):
    """p906"""
    fname = "Configure and Measure"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target Name", targetCollection, group, name)
    sdk.SetStringArg("Measurement Mode", profileName)
    sdk.SetBoolArg("Measure Immediately", measureImmediately)
    sdk.SetBoolArg("Wait for Completion", waitForCompletion)
    sdk.SetDoubleArg("Timeout in Seconds", timeoutInSecs)
    sdk.ExecuteStep()
    result = sdk.GetMPStepResult(0)
    return result


def compute_CTE_scale_factor(cte, parttemp):
    """p929"""
    fname = "Compute CTE Scale Factor"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    sdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    sdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    sdk.ExecuteStep()
    scaleFactor = sdk.GetDoubleArg("Scale Factor", 0.0)
    return scaleFactor


def set_instrument_scale_absolute(collection, instid, scaleFactor):
    """p931"""
    fname = "Set (absolute) Instrument Scale Factor (CAUTION!)"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", collection, instid)
    sdk.SetDoubleArg("Scale Factor", scaleFactor)
    sdk.ExecuteStep()
    getResult(fname)


# ################################
# Chapter 13 - Robot Operations ##
# ################################


# ##################################
# Chapter 14 - Utility Operations ##
# ##################################
def  move_collection_to_folder(collection, folder):
    """p1055"""
    fname = 'Move Collection to Folder'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection", collection)
    sdk.SetStringArg("Folder Path", folder)
    sdk.ExecuteStep()
    getResult(fname)


def set_collection_notes(collection, notes):
    """p1063"""
    fname = 'Set Collection Notes'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection", collection)
    stringList = [notes]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    sdk.SetEditTextArg("Notes", vStringList)
    sdk.SetBoolArg("Append? (FALSE = Overwrite)", True)
    sdk.ExecuteStep( )
    getResult(fname)


def set_working_frame(col, name):
    """p1080"""
    fname = 'Set Working Frame'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("New Working Frame Name", col, name)
    sdk.ExecuteStep()
    getResult(fname)


def delete_objects(col, name, objtype):
    """p1089"""
    fname = 'Delete Objects'
    fprint(fname)
    sdk.SetStep(fname)
    objNameList = [f'{col}::{name}::{objtype}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Object Names", vObjectList)
    sdk.ExecuteStep()
    getResult(fname)


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