"""
SAPyLib = Spatial Analyzer Python .NET library

This library provides the communication layer to the SA SDK .NET dll.
Author: L. Ververgaard
Date: 2021-06-26
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

# SA SDK dll
sa_sdk_dll_file = os.path.join(basepath, r"dll\Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
NrkSdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = NrkSdk.Connect("127.0.0.1")
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

def getResult_(fname):
    """Get the methods execution result."""
    # result = getIntRef()
    boolean, result = NrkSdk.GetMPStepResult(0)
    dprint(f'{fname}: {boolean}, {result}')
    if result == -1:
        # SDKERROR = -1
        raise SystemError('Execution raised: SDKERROR!')
    elif result == 0:
        # UNDONE = 0
        raise IOError('Execution: undone.')
    elif result == 1:
        # INPROGRESS = 1
        dprint('Execution: inprogress.')
        return True
    elif result == 2:
        # DONESUCCESS = 2
        return True
    elif result == 3:
        # DONEFATALERROR = 3
        raise IOError('Execution: FAILED!')
    elif result == 4:
        # DONEMINORERROR = 4
        dprint('Execution: FAILED - minor error!')
        return True
    elif result == 5:
        # CURRENTTASK = 5
        raise IOError('Execution: current task')
    elif result == 6:
        # UNKNOWN = 6
        raise SystemError('I have no clue!')

def getResult(fname):
    """Get the methods execution result."""
    boolean, result = NrkSdk.GetMPStepResult(0)
    dprint(f'{fname}: {boolean}, {result}')
    if result == -1:
        # SDKERROR = -1
        raise SystemError('Execution raised: SDKERROR!')
    elif result == 0:
        # UNDONE = 0
        dprint('Execution: undone.')
        return True
    elif result == 1:
        # INPROGRESS = 1
        dprint('Execution: inprogress.')
        return True
    elif result == 2:
        # DONESUCCESS = 2
        return True
    elif result == 3:
        # DONEFATALERROR = 3
        dprint('Execution: FAILED!')
        return False
    elif result == 4:
        # DONEMINORERROR = 4
        dprint('Execution: FAILED - minor error!')
        return True
    elif result == 5:
        # CURRENTTASK = 5
        dprint('Execution: current task')
        return True
    elif result == 6:
        # UNKNOWN = 6
        raise SystemError('I have no clue!')

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
        self.collection = point[0]
        self.group = point[1]
        self.name = point[2]
        
        # Get XYZ
        pData = get_point_coordinate(self.collection, self.group, self.name)
        dprint(f'pData: {pData}')
        self.X = pData[0]
        self.Y = pData[1]
        self.Z = pData[2]


class CTE:
    AluminumCTE_1DegF = 0.0000131
    SteelCTE_1DegF = 0.0000065
    CarbonFiberCTE_1DegF = 0


# ##############################
# Chapter 2 - File Operations ##
# ##############################
def find_files_in_directory(directory, searchPattern):
    """p27"""
    fname = 'Find Files in Directory'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Directory", directory)
    NrkSdk.SetStringArg("File Name Pattern", searchPattern)
    NrkSdk.SetBoolArg("Recursive?", False)
    NrkSdk.ExecuteStep()
    results = getResult(fname)
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
    fname = ''
    fprint(fname)
    NrkSdk.SetStep(fname)

    NrkSdk.ExecuteStep()
    getResult(fname)


def ask_for_string(question, *args, **kwargs):
    """p117"""

    if 'initialanswer' in kwargs:
        initialanswer = kwargs['initialanswer']
    else:
        initialanswer = ''
    
    fname = 'Ask for String'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Question to ask", question)
    NrkSdk.SetBoolArg("Password Entry?", False)
    NrkSdk.SetStringArg("Initial Answer", initialanswer)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(fname)
    answer = NrkSdk.GetStringArg("Answer", '')
    return answer[1]


def ask_for_string_pulldown(question, answers):
    """p118"""
    fname = 'Ask for String (Pull-Down Version)'
    fprint(fname)
    NrkSdk.SetStep(fname)
    # question section
    questionList = [question, ]
    QvStringList = System.Runtime.InteropServices.VariantWrapper(questionList)
    NrkSdk.SetStringRefListArg("Question or Statement", QvStringList)
    # anwsers section
    AvStringList = System.Runtime.InteropServices.VariantWrapper(answers)
    NrkSdk.SetStringRefListArg("Possible Answers", AvStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(fname)
    answer = NrkSdk.GetStringArg("Answer", '')
    return answer[1]


def ask_for_user_decision_extended(question, choices):
    """p122"""
    fname = 'Ask for User Decision Extended'
    fprint(fname)
    NrkSdk.SetStep(fname)
    stringList = [question, ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Question or Statement", vStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    NrkSdk.ExecuteStep()
    getResult(fname)
    selected_answer = 0
    return selected_answer


# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################


# ###########################
# Chapter 5 - View Control ##
# ###########################
def hide_objects(collection, name, objtype):
    """p153"""
    fname = 'Hide Objects'
    fprint(fname)
    NrkSdk.SetStep(fname)
    objNameList = [f'{collection}::{name}::{objtype}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Hide", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(fname)


def show_hide_objects(collection, objtype, hide):
    """p154"""
    # Hide = hide/True
    # Show = hide/False
    fname = 'Show / Hide by Object Type'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetBoolArg("All Collections?", False)
    NrkSdk.SetCollectionNameArg("Specific Collection", collection)
    # Available options: 
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh", 
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group", 
    NrkSdk.SetObjectTypeArg("Object Type To Show / Hide", objtype)
    dprint(f'\t{collection}::{objtype} Hide? {hide}')
    NrkSdk.SetBoolArg("Hide? (Show = FALSE)", hide)
    NrkSdk.ExecuteStep()
    getResult(fname)


def show_objects(collection, objects, name):
    """p155"""
    fname = 'Show Objects'
    fprint(fname)
    NrkSdk.SetStep(fname)
    objNameList = [f'{collection}::{objects}::{name}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Objects To Show", vObjectList)
    NrkSdk.ExecuteStep()
    getResult(fname)


def show_hide_callout_view(collection, calloutname, show):
    """p159"""
    # Show = show/True
    # Hide = show/False
    fname = 'Show / Hide Callout View'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Callout View To Show", collection, calloutname)
    NrkSdk.SetBoolArg("Show Callout View?", show)
    NrkSdk.ExecuteStep()
    getResult(fname)


def hide_all_callout_views():
    """p160"""
    fname = 'Hide All Callout Views'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.ExecuteStep()
    getResult(fname)


def center_graphics_about_objects(ColWild, ObjWild):
    """p170"""
    fname = ''
    fprint(fname)
    NrkSdk.SetStep(fname)

    NrkSdk.ExecuteStep()
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
    NrkSdk.SetStep(fname)
    NrkSdk.SetPointNameArg("Original Point Name", orgCol, orgGrp, orgName)
    NrkSdk.SetPointNameArg("New Point Name", newCol, newGrp, newName)
    NrkSdk.SetBoolArg("Overwrite if exists?", False)
    NrkSdk.ExecuteStep()
    getResult(fname)


def rename_collection(fromName, toName):
    """p194"""
    fname = "Rename Collection"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Original Collection Name", fromName)
    NrkSdk.SetCollectionNameArg("New Collection Name", toName)
    NrkSdk.ExecuteStep()
    getResult(fname)


def delete_points(collection, group, point):
    """p197"""
    fname = 'Delete Points'
    fprint(fname)
    NrkSdk.SetStep(fname)
    ptNameList = [f'{collection}::{group}::{point}', ]
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg('Point Names', vPointObjectList)
    NrkSdk.ExecuteStep()
    getResult(fname)


def delete_points_wildcard_selection(Col, Group, ObjType, pName):
    """p198"""
    fname = "Delete Points WildCard Selection"
    fprint(fname)
    NrkSdk.SetStep(fname)
    objNameList = [f'{Col}::{Group}::{ObjType}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Groups to Delete From", vObjectList)
    NrkSdk.SetPointNameArg("WildCard Selection Names", "*", "*", pName)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_objects_from_surface_faces_runtime_select(*args, **kwargs):
    """p199"""
    fname = 'Construct Objects From Surface Faces - Runtime Select'
    fprint(fname)
    NrkSdk.SetStep(fname)
    if 'facetype' in kwargs:
        facetype = kwargs['facetype']
        if not facetype:
            dprint('\tNo facetype selected!')
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
    NrkSdk.SetBoolArg("Construct Planes?", state[0])
    NrkSdk.SetBoolArg("Construct Cylinders?", state[1])
    NrkSdk.SetBoolArg("Construct Spheres?", state[2])
    NrkSdk.SetBoolArg("Construct Cones?", state[3])
    NrkSdk.SetBoolArg("Construct Lines?", state[4])
    NrkSdk.SetBoolArg("Construct Points?", state[5])
    NrkSdk.SetBoolArg("Construct Circles?", state[6])
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_or_construct_default_collection(colName):
    """p201"""
    fname = "Set (or construct) default collection"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Collection Name", colName)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_collection(name, makeDefault):
    """p202"""
    # default collection = makeDefault/True
    # collection = makeDefault/False
    fname = "Construct Collection"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Collection Name", name)
    NrkSdk.SetStringArg("Folder Path", "")
    NrkSdk.SetBoolArg("Make Default Collection?", makeDefault)
    NrkSdk.ExecuteStep()
    getResult(fname)


def get_active_collection():
    """p203"""
    fname = 'Get Active Collection Name'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.ExecuteStep()
    getResult(fname)
    sValue = NrkSdk.GetCollectionNameArg("Currently Active Collection Name", '')
    if sValue[0]:
        return sValue[1]
    else:
        return False


def delete_collection(collection):
    """p204"""
    fname = "Delete Collection"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_a_point(collection, group, name, x, y, z):
    """p208"""
    fname = "Construct a Point in Working Coordinates"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetPointNameArg("Point Name", collection, group, name)
    NrkSdk.SetVectorArg("Working Coordinates", x, y, z)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_point_at_intersection_of_plane_and_line(collectionplane, plane,
                                                      collectionline, line,
                                                      collectionpoint, grouppoint, point
                                                      ):
    """p218"""
    fname = 'Construct Point at Intersection of Plane and Line'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", collectionplane, plane)
    NrkSdk.SetCollectionObjectNameArg("Line Name", collectionline, line)
    NrkSdk.SetPointNameArg("Resulting Point Name", collectionpoint, grouppoint, point)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_line_2_points(name, fCol, fGrp, fTarg, sCol, sGrp, sTarg):
    """p262"""
    fname = 'Construct Line 2 Points'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Line Name", "", name)
    NrkSdk.SetPointNameArg("First Point", fCol, fGrp, fTarg)
    NrkSdk.SetPointNameArg("Second Point", sCol, sGrp, sTarg)
    NrkSdk.ExecuteStep()
    getResult(fname)


def construct_plane(planeCol, planeName):
    """p271"""
    fname = 'Construct Plane'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Plane Name", planeCol, planeName)
    NrkSdk.SetVectorArg("Plane Center (in working coordinates)", 0.0, 0.0, 0.0)
    NrkSdk.SetVectorArg("Plane Normal (in working coordinates)", 0.0, 0.0, 1.0)
    NrkSdk.SetDoubleArg("Plane Edge Dimension", 0.0)
    NrkSdk.ExecuteStep()
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
    NrkSdk.SetStep(fname)
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
    getResult(fname)


def create_relationship_callout(collection, calloutname, relationshipcollection, relationshipname, *args, **kwargs):
    """p363"""
    fname = 'Create Relationship Callout'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection, calloutname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relationshipcollection, relationshipname)
    # get X position
    if 'xpos' in kwargs:
        xpos = kwargs['xpos']
    else:
        xpos=0.1
    # get Y position
    if 'ypos' in kwargs:
        ypos = kwargs['ypos']
    else:
        ypos = 0.1
    # set values
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    vStringList = System.Runtime.InteropServices.VariantWrapper([])
    NrkSdk.SetEditTextArg("Additional Notes (blank for none)", vStringList)
    NrkSdk.ExecuteStep()
    getResult(fname)


def create_text_callout(collection, calloutname, text, *args, **kwargs):
    """p365"""
    fname = 'Create Text Callout'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Destination Callout View", collection, calloutname)
    stringList = [text, ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Text", vStringList)
    # get X position
    if 'xpos' in kwargs:
        xpos = kwargs['xpos']
    else:
        xpos=0.1
    # get Y position
    if 'ypos' in kwargs:
        ypos = kwargs['ypos']
    else:
        ypos = 0.1
    # set values
    NrkSdk.SetDoubleArg("View X Position", xpos)
    NrkSdk.SetDoubleArg("View Y Position", ypos)
    NrkSdk.SetPointNameArg("Callout Anchor Point (Optional)", "", "", "")
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_default_callout_view_properties(calloutname):
    """p372"""
    fname = 'Set Default Callout View Properties'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Default Callout View Name", calloutname)
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
    getResult(fname)


def make_a_point_name_runtime_select(txt):
    """p398"""
    fname = 'Make a Point Name - Runtime Select'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("User Prompt", txt)
    NrkSdk.ExecuteStep()
    getResult(fname)
    point = NrkSdk.GetPointNameArg("Resultant Point Name", '', '', '')
    dprint(f'\tPoint: {point}')
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
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Group Name", collection, group)
    NrkSdk.ExecuteStep()
    getResult(fname)
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if ptList[0]:
        points = []
        for i in range(ptList[1].GetLength(0)):
            points.append(ptList[1][i].split('::'))
        return points
    else:
        return False


def make_a_point_name_ref_list_runtime_select(prompt):
    """p402"""
    fname = 'Make a Point Name Ref List - Runtime Select'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("User Prompt", prompt)
    NrkSdk.ExecuteStep( )
    getResult(fname)
    userPtList = System.Runtime.InteropServices.VariantWrapper([])
    ptList = NrkSdk.GetPointNameRefListArg("Resultant Point Name List", userPtList)
    if ptList[0]:
        points = []
        for i in range(ptList[1].GetLength(0)):
            points.append(ptList[1][i].split('::'))
        return points
    else:
        return False


def make_a_collection_name_runtime_select(txt):
    """p408"""
    fname = 'Make a Collection Name - Runtime Select'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("User Prompt", txt)
    NrkSdk.ExecuteStep()
    getResult(fname)
    sValue = NrkSdk.GetCollectionNameArg("Resultant Collection Name", '')
    if sValue[0]:
        return sValue[1]
    else:
        return False


def make_a_collection_object_name_ref_list_by_type(collection, objtype):
    """p416"""
    fname = 'Make a Collection Object Name Ref List - By Type'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Collection", collection)
    # Available options: 
    # "Any", "B-Spline", "Circle", "Cloud", "Scan Stripe Cloud", 
    # "Cross Section Cloud", "Cone", "Cylinder", "Datum", "Ellipse", 
    # "Frame", "Frame Set", "Line", "Paraboloid", "Perimeter", 
    # "Plane", "Point Group", "Point Set", "Poly Surface", "Scan Stripe Mesh", 
    # "Slot", "Sphere", "Surface", "Torus", "Vector Group", 
    NrkSdk.SetObjectTypeArg("Object Type", objtype)
    NrkSdk.ExecuteStep()
    getResult(fname)
    userObjectList = System.Runtime.InteropServices.VariantWrapper([])
    objectList = NrkSdk.GetCollectionObjectNameRefListArg("Resultant Collection Object Name List", userObjectList)
    dprint(f'\tobjectList: {objectList}')
    dprint(f'\tobjectList[0]: {objectList[0]}')
    dprint(f'\tobjectList[1]: {objectList[1]}')
    dprint(f'\tLength objectList[1]: {objectList[1].GetLength(0)}')
    if objectList[0]:
        objects = []
        for i in range(objectList[1].GetLength(0)):
            objects.append(objectList[1][i].split('::'))
        return objects
    else:
        return False


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def get_number_of_collections():
    """p465"""
    fname = 'Get Number of Collections'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.ExecuteStep()
    getResult(fname)
    n = NrkSdk.GetIntegerArg('Total Count', 0)
    dprint(f'n: {n}')
    if n[0]:
        return n[1]
    else:
        return False


def get_ith_collection_name(i):
    """p466"""
    fname = 'Get i-th Collection Name'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetIntegerArg('Collection Index', i)
    NrkSdk.ExecuteStep()
    getResult(fname)
    collection = NrkSdk.GetCollectionNameArg('Resultant Name', '')
    dprint(f'\tCollection: {collection}')
    if collection[0]:
        return collection[1]
    else:
        return False


def set_vector_group_colorization_options_selected(collection, vectorgroup, *args, **kwargs):
    """p485"""
    fname = 'Set Vector Group Colorization Options (Selected)'
    fprint(fname)
    NrkSdk.SetStep("Set Vector Group Colorization Options (Selected)")
    vgNameList = [f'{collection}::{vectorgroup}', ]
    vVectorGroupNameList = System.Runtime.InteropServices.VariantWrapper(vgNameList)
    NrkSdk.SetCollectionVectorGroupNameRefListArg("Vector Groups to be Set", vVectorGroupNameList)
    NrkSdk.SetColorizationOptionsArg("Colorization Options",
                                     kwargs['color_profile'],
                                     kwargs['base_high_color'],
                                     kwargs['base_mid_color'],
                                     kwargs['base_low_color'],
                                     kwargs['tubes'],
                                     kwargs['arrows'],
                                     kwargs['show_labels'],
                                     kwargs['vector_magnification'],
                                     kwargs['vector_width'],
                                     kwargs['blotches'],
                                     kwargs['blotch_size'],
                                     kwargs['show_out_of_tol'],
                                     kwargs['color_bar'],
                                     kwargs['show_percentages'],
                                     kwargs['show_fractions'],
                                     kwargs['x1'],
                                     kwargs['x2'],
                                     kwargs['high_tol'],
                                     kwargs['low_tol'],
                                     )
    NrkSdk.ExecuteStep()
    getResult(fname)


def get_point_coordinate(collection, group, pointname):
    """p489"""
    fname = 'Get Point Coordinate'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetPointNameArg("Point Name", collection, group, pointname)
    NrkSdk.ExecuteStep()
    if not getResult(fname):
        return False
    Vector = NrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    dprint(f'\tVector: {Vector}')
    if Vector[0]:
        xVal = Vector[1]
        yVal = Vector[2]
        zVal = Vector[3]
        return (xVal, yVal, zVal)
    else:
        return False


def get_point_to_point_distance(p1col, p1grp, p1name, p2col, p2grp, p2name):
    """p495"""
    fname = 'Get Point To Point Distance'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetPointNameArg("First Point", p1col, p1grp, p1name)
    NrkSdk.SetPointNameArg("Second Point", p2col, p2grp, p2name)
    NrkSdk.ExecuteStep()
    getResult(fname)
    # xVal, yVal, zValNrkSdk.GetVectorArg("Vector Representation", 0.0, 0.0, 0.0)
    xval = NrkSdk.GetDoubleArg("X Value", 0.0)
    yval = NrkSdk.GetDoubleArg("Y Value", 0.0)
    zval = NrkSdk.GetDoubleArg("Z Value", 0.0)
    mag = NrkSdk.GetDoubleArg("Magnitude", 0.0)
    returndict = {'xval': xval[1],
                  'yval': yval[1],
                  'zval': zval[1],
                  'mag': mag[1],
                  }
    return returndict


def set_default_colorization_options():
    """p531"""
    fname = 'Set Default Colorization Options'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColorizationOptionsArg("Colorization Options", "Continuous", "Blue", "Green", "Red", False, True, False, 100.0, 1, False, 0.1, False, False, True, False, 0.5, -0.5, 0.03, -0.5)
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_vector_group_display_attributes(magnification, blotch, tolerance):
    """p532"""
    fname = 'Set Vector Group Display Attributes'
    fprint(fname)
    NrkSdk.SetStep(fname)
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
    getResult(fname)


def fit_geometry_to_point_group(geomType,
                                dataCol, dataGroup,
                                resultCol, resultName,
                                profilename, reportdiv, fittol, outtol
                                ):
    """p547"""
    fname = 'Fit Geometry to Point Group'
    fprint(fname)
    NrkSdk.SetStep(fname)
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
    NrkSdk.ExecuteStep( )
    getResult(fname)


def best_fit_group_to_group(refCollection, refGroup, 
                            corCollection, corGroup, 
                            showDialog, rmsTol, maxTol, allowScale):
    """p551"""
    fname = "Best Fit Transformation - Group to Group"
    fprint(fname)
    NrkSdk.SetStep(fname)
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
    NrkSdk.SetFilePathArg("File Path for CSV Text Report (requires Show Interface = TRUE)",
                       "", False)
    NrkSdk.ExecuteStep()
    r = getResult(fname)
    return r


def make_point_to_point_relationship(relCol, relName, p1col, p1grp, p1, p2col, p2grp, p2):
    """p632"""
    fname = 'Make Point to Point Relationship'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetPointNameArg("First Point Name", p1col, p1grp, p1)
    NrkSdk.SetPointNameArg("Second Point Name", p2col, p2grp, p2)
    NrkSdk.SetToleranceVectorOptionsArg("Tolerance",
                                        False, 0.0, False, 0.0, False, 0.0, False, 0.0,
                                        False, 0.0, False, 0.0, False, 0.0, False, 0.0)
    NrkSdk.SetToleranceVectorOptionsArg("Constraint",
                                        True, 0.0, True, 0.0, True, 0.0, False, 0.0,
                                        True, 0.0, True, 0.0, True, 0.0, False, 0.0)
    NrkSdk.ExecuteStep()
    getResult(fname)


def make_group_to_nominal_group_relationship(relCol, relName,
                                             nomCol, nomGrp,
                                             meaCol, meaGrp,
                                             ):
    """p646"""
    fname = 'Make Group to Nominal Group Relationship'
    fprint(fname)
    NrkSdk.SetStep(fname)
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
        False, 0.0, False, 0.0, False, 0.0, True, 1.5,
        False, 0.0, False, 0.0, False, 0.0, True, 0.0
        )
    NrkSdk.SetToleranceVectorOptionsArg(
        "Constraint",
        True, 0.0, True, 0.0, True, 0.0, False, 0.0,
        True, 0.0, True, 0.0, True, 0.0, False, 0.0
        )
    NrkSdk.SetDoubleArg("Fit Weight", 1.0)
    NrkSdk.ExecuteStep()
    getResult(fname)


def make_geometry_fit_and_compare_to_nominal_relationship(relatCol, relatName,
                                                          nomCol, nomName,
                                                          data,
                                                          resultCol, resultName
                                                          ):
    """p649"""
    fname = 'Make Geometry Fit and Compare to Nominal Relationship'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relatCol, relatName)
    NrkSdk.SetCollectionObjectNameArg("Nominal Geometry", nomCol, nomName)
    objNameList = []
    for item in data:
        objNameList.append(f'{item[0]}::{item[1]}::Point Group')
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Point Groups to Fit", vObjectList)
    NrkSdk.SetCollectionObjectNameArg("Resulting Object Name (Optional)", resultCol, resultName)
    NrkSdk.SetStringArg("Fit Profile Name (Optional)", "")
    NrkSdk.ExecuteStep()
    getResult(fname)


def get_geom_relationship_criteria(relCol, relName, criteria):
    """p660"""
    fname = 'Get Geom Relationship Criteria'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetStringArg("Criteria", criteria)
    NrkSdk.ExecuteStep()
    if not getResult(fname):
        return False
    nominal = NrkSdk.GetDoubleArg("Nominal", 0.0)
    measured = NrkSdk.GetDoubleArg("Measured", 0.0)
    delta = NrkSdk.GetDoubleArg("Delta", 0.0)
    lowtol = NrkSdk.GetDoubleArg("Low Tolerance", 0.0)
    hightol = NrkSdk.GetDoubleArg("High Tolerance", 0.0)
    deltaweight = NrkSdk.GetDoubleArg("Optimization: Delta Weight", 0.0)
    outoftolweight = NrkSdk.GetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    dict_return = {'nominal': nominal[1],
                   'measured': measured[1],
                   'delta': delta[1],
                   'lowtol': lowtol[1],
                   'hightol': hightol[1],
                   'deltaweight': deltaweight[1],
                   'outoftolweight': outoftolweight[1],
                   }
    # print(dict_return)
    return dict_return


def set_relationship_associated_data(relCol, group, collectionmeasured, groupmeasured):
    """p664"""
    # The function only excepts 'groups' as an input.
    # Individual points, point clouds and objects aren't support for now.
    fname = 'Set Relationship Associated Data'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, group)
    # individual points
    ptNameList = []
    vPointObjectList = System.Runtime.InteropServices.VariantWrapper(ptNameList)
    NrkSdk.SetPointNameRefListArg("Individual Points", vPointObjectList)
    # point groups
    objNameList = [f'{collectionmeasured}::{groupmeasured}::Point Group', ]
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
    getResult(fname)


def set_relationship_reporting_frame(relCol, relGrp, frmCol, frmName):
    """p679"""
    fname = 'Set Relationship Reporting Frame'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    NrkSdk.SetCollectionObjectNameArg("Reporting Frame", frmCol, frmName)
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_geom_relationship_criteria(relCol, relName, criteriatype):
    """p680"""
    fname = 'Set Geom Relationship Criteria'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    # Flatness
    if criteriatype == 'Flatness':
        NrkSdk.SetStringArg("Criteria", "Flatness")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.1, True, -0.1)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Centroid Z
    elif criteriatype == 'Centroid Z':
        NrkSdk.SetStringArg("Criteria", "Centroid Z")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Avg Distance Between
    elif criteriatype == 'Avg Dist Between':
        NrkSdk.SetStringArg("Criteria", "Avg Dist Between")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 3.0, True, -3.0)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    # Length
    elif criteriatype == 'Length':
        NrkSdk.SetStringArg("Criteria", "Length")
        NrkSdk.SetBoolArg("Show in Report", True)
        NrkSdk.SetToleranceScalarOptionsArg("Tolerance Options", True, 0.05, True, -0.05)
        NrkSdk.SetDoubleArg("Optimization: Delta Weight", 0.0)
        NrkSdk.SetDoubleArg("Optimization: Out of Tolerance Weight", 0.0)
    else:
        dprint('\tIncorrect criteria type set!')
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_geom_relationship_auto_vectors_nominal_avn(collection, relationshipname, bool):
    """p690"""
    fname = 'Set Geom Relationship Auto Vectors Nominal (AVN)'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationshipname)
    NrkSdk.SetBoolArg("Create Auto Vectors AVN", bool)
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_relationship_auto_vectors_fit_avf(collection, relationshipname, bool):
    """p691"""
    fname = 'Set Relationship Auto Vectors Fit (AVF)'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", collection, relationshipname)
    NrkSdk.SetBoolArg("Create Auto Vectors AVF", bool)
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_relationship_desired_meas_count(relCol, relName, count):
    """p693"""
    fname = 'Set Relationship Desired Meas Count'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetIntegerArg("Desired Measurement Count", count)
    NrkSdk.ExecuteStep()
    getResult(fname)


def set_relationship_tolerance_vector_type(relCol, relName):
    """p697"""
    fname = 'Set Relationship Tolerance (Vector Type)'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relName)
    NrkSdk.SetToleranceVectorOptionsArg("Vector Tolerance",
                                        False, 0.000000, False, 0.000000, False, 0.000000, True, 1.0,
                                        False, 0.000000, False, 0.000000, False, 0.000000, True, 0.0)
    NrkSdk.ExecuteStep()
    getResult(fname)


# ###################################
# Chapter 9 - Reporting Operations ##
# ###################################
def set_relationship_report_options(relCol, relGrp, *args, **kwargs):
    """p770"""
    fname = 'Set Relationship Report Options'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("Relationship Name", relCol, relGrp)
    NrkSdk.SetPointDeltaReportOptionsArg("Report Options",
                                         kwargs['coord_system'],
                                         kwargs['record_format'],
                                         kwargs['a'],
                                         kwargs['b'],
                                         kwargs['c'],
                                         kwargs['d'],
                                         kwargs['e'],
                                         kwargs['f'],
                                         kwargs['summary'],
                                         kwargs['h'],
                                         kwargs['i'],
                                         kwargs['j'],
                                         )
    NrkSdk.ExecuteStep()
    getResult(fname)


def notify_user_text_array(txt, timeout=0):
    """p804"""
    fname = 'Notify User Text Array'
    fprint(fname)
    NrkSdk.SetStep(fname)
    stringList = [txt, ]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Notification Text", vStringList)
    NrkSdk.SetFontTypeArg("Font", "MS Shell Dlg", 8, 0, 0, 0)
    NrkSdk.SetBoolArg("Auto expand to fit text?", False)
    NrkSdk.SetIntegerArg("Display Timeout", timeout)
    NrkSdk.ExecuteStep( )
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
    NrkSdk.SetStep(fname)
    NrkSdk.ExecuteStep()
    getResult(fname)
    # instId = NrkSdk.GetIntegerArg("Instrument ID", 0)
    # dprint(f'\tinstId: {instId}')
    ColInstID = NrkSdk.GetColInstIdArg("Instrument ID", '', 0)
    dprint(f'\tColInstID: {ColInstID}')
    return (ColInstID[1], ColInstID[2])


def point_at_target(instCol, instId, collection, group, name):
    """p877"""
    fname = "Point At Target"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColInstIdArg("Instrument ID", instCol, instId)
    NrkSdk.SetPointNameArg("Target ID", collection, group, name)
    NrkSdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    NrkSdk.ExecuteStep()
    getResult(fname)


def add_new_instrument(instName):
    """p889"""
    fname = "Add New Instrument"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetInstTypeNameArg("Instrument Type", instName)
    NrkSdk.ExecuteStep()
    getResult(fname)
    ColInstID = NrkSdk.GetColInstIdArg("Instrument Added (result)", '', 0)
    dprint(f'\tColInstID: {ColInstID}')
    return ColInstID[1], ColInstID[2]


def start_instrument(InstCol, InstID, initialize, simulation):
    """p902"""
    fname = "Start Instrument Interface"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColInstIdArg("Instrument's ID", InstCol, InstID)
    NrkSdk.SetBoolArg("Initialize at Startup", initialize)
    NrkSdk.SetStringArg("Device IP Address (optional)", "")
    NrkSdk.SetIntegerArg("Interface Type (0=default)", 0)
    NrkSdk.SetBoolArg("Run in Simulation", simulation)
    NrkSdk.ExecuteStep()
    getResult(fname)


def stop_instrument(collection, instid):
    """p903"""
    fname = "Stop Instrument Interface"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection, instid)
    NrkSdk.ExecuteStep()
    getResult(fname)


def configure_and_measure(instrumentCollection, instId, 
                          targetCollection, group, name, 
                          profileName, measureImmediately, 
                          waitForCompletion, timeoutInSecs):
    """p906"""
    fname = "Configure and Measure"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColInstIdArg("Instrument's ID", instrumentCollection, instId)
    NrkSdk.SetPointNameArg("Target Name", targetCollection, group, name)
    NrkSdk.SetStringArg("Measurement Mode", profileName)
    NrkSdk.SetBoolArg("Measure Immediately", measureImmediately)
    NrkSdk.SetBoolArg("Wait for Completion", waitForCompletion)
    NrkSdk.SetDoubleArg("Timeout in Seconds", timeoutInSecs)
    NrkSdk.ExecuteStep()
    result = getResult(fname)
    return result


def compute_CTE_scale_factor(cte, parttemp):
    """p929"""
    fname = "Compute CTE Scale Factor"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    NrkSdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    NrkSdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    NrkSdk.ExecuteStep()
    getResult(fname)
    scaleFactor = NrkSdk.GetDoubleArg("Scale Factor", 0.0)
    if scaleFactor[0]:
        return scaleFactor[1]
    else:
        return False


def set_instrument_scale_absolute(collection, instid, scaleFactor):
    """p931"""
    fname = "Set (absolute) Instrument Scale Factor (CAUTION!)"
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetColInstIdArg("Instrument's ID", collection, instid)
    NrkSdk.SetDoubleArg("Scale Factor", scaleFactor)
    NrkSdk.ExecuteStep()
    getResult(fname)


# ################################
# Chapter 13 - Robot Operations ##
# ################################


# ##################################
# Chapter 14 - Utility Operations ##
# ##################################
def delete_folder(foldername):
    """p1054"""
    fname = 'Delete Folder'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Folder Path", foldername)
    NrkSdk.ExecuteStep()
    getResult(fname)


def move_collection_to_folder(collection, folder):
    """p1055"""
    fname = 'Move Collection to Folder'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(fname)


def get_folders_by_wildcard(search):
    """p1057"""
    fname = 'Get Folders by Wildcard'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Search String", search)
    NrkSdk.SetBoolArg("Case Sensitive Search", False)
    NrkSdk.ExecuteStep()
    getResult(fname)
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
    fname = 'Get Folder Collections'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetStringArg("Folder Path", folder)
    NrkSdk.ExecuteStep()
    getResult(fname)
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
    fname = 'Set Collection Notes'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionNameArg("Collection", collection)
    stringList = [notes]
    vStringList = System.Runtime.InteropServices.VariantWrapper(stringList)
    NrkSdk.SetEditTextArg("Notes", vStringList)
    NrkSdk.SetBoolArg("Append? (FALSE = Overwrite)", True)
    NrkSdk.ExecuteStep( )
    getResult(fname)


def set_working_frame(col, name):
    """p1080"""
    fname = 'Set Working Frame'
    fprint(fname)
    NrkSdk.SetStep(fname)
    NrkSdk.SetCollectionObjectNameArg("New Working Frame Name", col, name)
    NrkSdk.ExecuteStep()
    getResult(fname)


def delete_objects(col, name, objtype):
    """p1089"""
    fname = 'Delete Objects'
    fprint(fname)
    NrkSdk.SetStep(fname)
    objNameList = [f'{col}::{name}::{objtype}', ]
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    NrkSdk.SetCollectionObjectNameRefListArg("Object Names", vObjectList)
    NrkSdk.ExecuteStep()
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