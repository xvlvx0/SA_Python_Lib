"""
SAPyLib = Spatial Analyzer Python .NET library

This library provides the communication layer to the SA SDK .NET dll.
Author: L. Ververgaard
Date: 2020-04-08
Version: 0
"""
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

clr.AddReference(os.path.join(basepath, r"dll\SAPyTools"))
from SAPyTools import MPHelper

# SA SDK dll
sa_sdk_dll_file = os.path.join(basepath, r"dll\Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
sdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = sdk.Connect("127.0.0.1")
if not SAConnected:
    raise IOError('Connection to SA SDK failed!')
    sys.exit(1)

# MP Helper
class MPResult:
    SDKERROR = -1
    UNDONE = 0
    INPROGRESS = 1
    DONESUCCESS = 2
    DONEFATALERROR = 3
    DONEMINORERROR = 4
    CURRENTTASK = 5
    UNKNOWN = 6
mphelper = MPHelper(sdk)
mpresult = MPResult()

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
        raise SystemError('Execution raised: SDKERROR!')
    elif result == 0:
        dprint('Execution: undone.')
        raise IOError()
    elif result == 1:
        dprint('Execution: inprogress.')
    elif result == 2:
        return True
    elif result == 3:
        dprint('Execution: FAILED!')
        raise IOError()
    elif result == 4:
        dprint('Execution: FAILED - minor error!')
        return True
    elif result == 5:
        dprint('Execution: current task')
        raise IOError()
    elif result == 6:
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


def delete_objects(listobj):
    mphelper.DeleteObjects(List[str](listobj))


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


# def collection_existence_test(ColName):
#     """p112"""
#     fname = 'Collection Existence Test'
#     fprint(fname)
#    sdk.SetStep(fname)
#     sdk.SetCollectionNameArg("Collection Name to check", ColName)
#     sdk.ExecuteStep()
#     r = getResult(fname)


def ask_for_string():
    """p117"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)
    sdk.ExecuteStep()
    if not getResult(fname):
        return False
    else:
        string = ''
        return string

def ask_for_string_pulldown(input):
    """p118"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)
    input = ['', '', '', ]
    sdk.ExecuteStep()
    if not getResult(fname):
        return False
    else:
        index = 0
        string = ''
        return index, string


# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################


# ###########################
# Chapter 5 - View Control ##
# ###########################
def hide_objects(ColObjNameRefList):
    """p153"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

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


def show_objects(ColObjNameRefList):
    """p155"""
    fname = ''
    fprint(fname)
    sdk.SetStep(fname)

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
    fname = ''
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
def rename_collection(fromName, toName):
    """p194"""
    fname = "Rename Collection"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Original Collection Name", fromName)
    sdk.SetCollectionNameArg("New Collection Name", toName)
    sdk.ExecuteStep()
    getResult(fname)


def delete_points_wildcard_selection(colObjNameRefList, pName):
    """p198"""
    fname = "Delete Points WildCard Selection"
    fprint(fname)
    sdk.SetStep(fname)
    objNameList = [colObjNameRefList, ]  # "Nominals::My Points::Point Group"
    vObjectList = System.Runtime.InteropServices.VariantWrapper(objNameList)
    sdk.SetCollectionObjectNameRefListArg("Groups to Delete From", vObjectList)
    sdk.SetPointNameArg("WildCard Selection Names", "*", "*", pName)
    sdk.ExecuteStep()
    getResult(fname)


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
    if not getResult(fname):
        return False
    else:
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


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def fit_geometry_to_point_group(geomType, dataCol, dataGroup, resultCol,
                                resultGroup, name, report, fit, oot
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
    sdk.SetCollectionObjectNameArg("Resulting Object Name", resultCol, resultGroup)
    sdk.SetStringArg("Fit Profile Name", name)
    sdk.SetBoolArg("Report Deviations", report)
    sdk.SetDoubleArg("Fit Interface Tolerance (-1.0 use profile)", fit)
    sdk.SetBoolArg("Ignore Out of Tolerance Points", oot)
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


def make_geometry_fit_and_compare_to_nominal_relationship(relatCol, relatName, nomCol,
                                                          nomName, data, resultCol, resultName
                                                          ):
    """p649"""
    fname = 'Make Geometry Fit and Compare to Nominal Relationship'
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetCollectionObjectNameArg("Relationship Name", relatCol, relatName)
    sdk.SetCollectionObjectNameArg("Nominal Geometry", nomCol, nomName)
    vObjectList = System.Runtime.InteropServices.VariantWrapper(data)
    sdk.SetCollectionObjectNameRefListArg("Point Groups to Fit", vObjectList)
    sdk.SetCollectionObjectNameArg("Resulting Object Name (Optional)", resultCol, resultName)
    sdk.SetStringArg("Fit Profile Name (Optional)", "")
    sdk.ExecuteStep()
    getResult(fname)


# ###################################
# Chapter 9 - Reporting Operations ##
# ###################################

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


def point_at_target(instrumentCollection, instId, collection, group, name):
    """p877"""
    fname = "Point At Target"
    fprint(fname)
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target ID", collection, group, name)
    sdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    sdk.ExecuteStep()
    fprint(fname)


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