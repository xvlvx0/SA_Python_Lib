import os
import clr

basepath = r"C:\Users\user\github\SAPythonDemo"

# Set the debugging reporting value
DEBUG = True

clr.AddReference("System")
clr.AddReference("System.Reflection")
import System
import System.Reflection
from System.Collections.Generic import List

clr.AddReference(os.path.join(basepath, r"SAPyTools\obj\Debug\SAPyTools"))
from SAPyTools import MPHelper

# SA SDK dll
sa_sdk_dll_file = os.path.join(basepath, r"SAPyTools\obj\Debug\Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
sdk = System.Activator.CreateInstance(sa_sdk_class_type)
SAConnected = sdk.Connect("127.0.0.1")

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

# #################
# Helper methods ##
# #################
def dprint(val):
    """Print a debugging value."""
    if DEBUG:
        print(f'DEBUG: {val}')


def getResult(fname):
    """Get the methods execution result."""
    result = 0
    sdk.GetMPStepResult(result)
    dprint(f'{fname}: {result}')
    return int(result)


def getIntRef():
    return 0


def getStrRef():
    return ''


def getDoubleRef():
    return 0.0


class Point3D:
    def __init__(self, x, y, z):
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

# ###############################
# Chapter 4 - MP Task Overview ##
# ###############################

# ###########################
# Chapter 5 - View Control ##
# ###########################

# ######################################
# Chapter 6 - Cloud Viewer Operations ##
# ######################################

# ######################################
# Chapter 7 - Construction Operations ##
# ######################################
def rename_collection(fromName, toName):
    """p194"""
    fname = "Rename Collection"
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Original Collection Name", fromName)
    sdk.SetCollectionNameArg("New Collection Name", toName)
    sdk.ExecuteStep()
    getResult(fname)


def delete_points_wildcard_selection(colObjNameRefList, pName):
    """p198"""
    fname = "Delete Points WildCard Selection"
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
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection Name", colName)
    sdk.ExecuteStep()
    getResult(fname)


def construct_collection(name, makeDefault):
    """p202"""
    fname = "Construct Collection"
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Collection Name", name)
    sdk.SetStringArg("Folder Path", "")
    sdk.SetBoolArg("Make Default Collection?", makeDefault)
    sdk.ExecuteStep()
    getResult(fname)


def delete_collection(collection):
    """p204"""
    fname = "Delete Collection"
    sdk.SetStep(fname)
    sdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    sdk.ExecuteStep()
    getResult(fname)


def construct_a_point(collection, group, name, x, y, z):
    """p208"""
    fname = "Construct a Point in Working Coordinates"
    sdk.SetStep(fname)
    sdk.SetPointNameArg("Point Name", collection, group, name)
    sdk.SetVectorArg("Working Coordinates", x, y, z)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


# ##################################
# Chapter 8 - Analysis Operations ##
# ##################################
def best_fit_group_to_group(refCollection, refGroup, 
                            corCollection, corGroup, 
                            showDialog, rmsTol, maxTol, allowScale):
    """p551"""
    fname = "Best Fit Transformation - Group to Group"
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
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


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
def point_at_target(instrumentCollection, instId, collection, group, name):
    """p877"""
    fname = "Point At Target"
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target ID", collection, group, name)
    sdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


def add_new_instrument(saname):
    """p889"""
    fname = "Add New Instrument"
    sdk.SetStep(fname)
    sdk.SetInstTypeNameArg("Instrument Type", saname)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)

    instid = getIntRef()
    collection = getStrRef()
    sdk.GetColInstIdArg("Instrument Added (result)", collection, instid)
    return instid, int(result)


def start_instrument(instid, initialize, simulation):
    """p902"""
    fname = "Start Instrument Interface"
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", "", instid)
    sdk.SetBoolArg("Initialize at Startup", initialize)
    sdk.SetStringArg("Device IP Address (optional)", "")
    sdk.SetIntegerArg("Interface Type (0=default)", 0)
    sdk.SetBoolArg("Run in Simulation", simulation)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


def stop_instrument(collection, instid):
    """p903"""
    fname = "Stop Instrument Interface"
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", collection, instid)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


def configure_and_measure(instrumentCollection, instId, 
                          targetCollection, group, name, 
                          profileName, measureImmediately, 
                          waitForCompletion, timeoutInSecs):
    """p906"""
    fname = "Configure and Measure"
    sdk.SetStep(fname)
    sdk.SetColInstIdArg("Instrument's ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target Name", targetCollection, group, name)
    sdk.SetStringArg("Measurement Mode", profileName)
    sdk.SetBoolArg("Measure Immediately", measureImmediately)
    sdk.SetBoolArg("Wait for Completion", waitForCompletion)
    sdk.SetDoubleArg("Timeout in Seconds", timeoutInSecs)
    sdk.ExecuteStep()
    result = getIntRef()
    sdk.GetMPStepResult(result)
    return int(result)


def compute_CTE_scale_factor(cte, parttemp):
    """p929"""
    fname = "Compute CTE Scale Factor"
    sdk.SetStep(fname)
    sdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    sdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    sdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    sdk.ExecuteStep()
    scaleFactor = getDoubleRef()
    sdk.GetDoubleArg("Scale Factor", scaleFactor)
    return scaleFactor


def set_instrument_scale_absolute(collection, instid, scaleFactor):
    """p931"""
    fname = "Set (absolute) Instrument Scale Factor (CAUTION!)"
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