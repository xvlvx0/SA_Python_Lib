import sys
import os
import clr

basepath = r"C:\Users\user\github\SAPythonDemo"

clr.AddReference("System")
clr.AddReference("System.Reflection")
import System
import System.Reflection
#from System.Collections.Generic import List

clr.AddReference(os.path.join(basepath, r"SAPyTools\obj\Debug\SAPyTools"))
from SAPyTools import MPHelper

# SA SDK dll
sa_sdk_dll_file = os.path.join(basepath, r"SAPyTools\obj\Debug\Interop.SpatialAnalyzerSDK.dll")
sa_sdk_dll = System.Reflection.Assembly.LoadFile(sa_sdk_dll_file)
sa_sdk_class_type = sa_sdk_dll.GetType("SpatialAnalyzerSDK.SpatialAnalyzerSDKClass")
sdk = System.Activator.CreateInstance(sa_sdk_class_type)

#sdk = System.Activator.CreateInstance(
#    type=System.Type.GetTypeFromProgID('SpatialAnalyzerSDK.Application'))

SAConnected = sdk.Connect("127.0.0.1")
mphelper = MPHelper(sdk)


class Point3D:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z


class MPResult:
    SDKERROR = -1
    UNDONE = 0
    INPROGRESS = 1
    DONESUCCESS = 2
    DONEFATALERROR = 3
    DONEMINORERROR = 4
    CURRENTTASK = 5
    UNKNOWN = 6

mpresult = MPResult()

class CTE:
    AluminumCTE_1DegF = 0.0000131
    SteelCTE_1DegF = 0.0000065
    CarbonFiberCTE_1DegF = 0

def getIntRef():
    return clr.Reference[int]()


def getStrRef():
    return clr.Reference[str]()

def getDoubleRef():
    return clr.Reference[float]()


def construct_collection(name, makeDefault):
    sdk.SetStep("Construct Collection")
    sdk.SetCollectionNameArg("Collection Name", name)
    sdk.SetStringArg("Folder Path", "")
    sdk.SetBoolArg("Make Default Collection?", makeDefault)
    sdk.ExecuteStep()


def delete_collection(collection):
    sdk.SetStep("Delete Collection")
    sdk.SetCollectionNameArg("Name of Collection to Delete", collection)
    sdk.ExecuteStep()


def set_or_construct_default_collection(colName):
    sdk.SetStep("Set (or construct) default collection")
    sdk.SetCollectionNameArg("Collection Name", colName)
    sdk.ExecuteStep()


def add_new_instrument(saname):
    result = getIntRef()
    sdk.SetStep("Add New Instrument")
    sdk.SetInstTypeNameArg("Instrument Type", saname)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    instid = getIntRef()
    collection = getStrRef()
    sdk.GetColInstIdArg("Instrument Added (result)", collection, instid)
    return instid, int(result)


def construct_a_point(collection, group, name, x, y, z):
    result = getIntRef()
    sdk.SetStep("Construct a Point in Working Coordinates")
    sdk.SetPointNameArg("Point Name", collection, group, name)
    sdk.SetVectorArg("Working Coordinates", x, y, z)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    return int(result)


def point_at_target(instrumentCollection, instId, collection, group, name):
    result = getIntRef()
    sdk.SetStep("Point At Target")
    sdk.SetColInstIdArg("Instrument ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target ID", collection, group, name)
    sdk.SetFilePathArg("HTML Prompt File (optional)", "", False)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    return int(result)


def start_instrument(instid, initialize, simulation):
    result = getIntRef()
    sdk.SetStep("Start Instrument Interface")
    sdk.SetColInstIdArg("Instrument's ID", "", instid)
    sdk.SetBoolArg("Initialize at Startup", initialize)
    sdk.SetStringArg("Device IP Address (optional)", "")
    sdk.SetIntegerArg("Interface Type (0=default)", 0)
    sdk.SetBoolArg("Run in Simulation", simulation)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    return int(result)


def stop_instrument(collection, instid):
    result = getIntRef()
    sdk.SetStep("Stop Instrument Interface")
    sdk.SetColInstIdArg("Instrument's ID", collection, instid)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    return int(result)


def configure_and_measure(instrumentCollection, instId, 
                          targetCollection, group, name, 
                          profileName, measureImmediately, 
                          waitForCompletion, timeoutInSecs):
    result = getIntRef()
    sdk.SetStep("Configure and Measure")
    sdk.SetColInstIdArg("Instrument's ID", instrumentCollection, instId)
    sdk.SetPointNameArg("Target Name", targetCollection, group, name)
    sdk.SetStringArg("Measurement Mode", profileName)
    sdk.SetBoolArg("Measure Immediately", measureImmediately)
    sdk.SetBoolArg("Wait for Completion", waitForCompletion)
    sdk.SetDoubleArg("Timeout in Seconds", timeoutInSecs)
    sdk.ExecuteStep()
    sdk.GetMPStepResult(result)
    return int(result)


def best_fit_group_to_group(refCollection, refGroup, 
                            corCollection, corGroup, 
                            showDialog, rmsTol, maxTol, allowScale):
    result = getIntRef()
    sdk.SetStep("Best Fit Transformation - Group to Group")
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
    sdk.GetMPStepResult(result)
    return int(result)

 
def delete_objects(listobj):
    mphelper.DeleteObjects(List[str](listobj))

def rename_collection(fromName, toName):
    sdk.SetStep("Rename Collection")
    sdk.SetCollectionNameArg("Original Collection Name", fromName)
    sdk.SetCollectionNameArg("New Collection Name", toName)
    sdk.ExecuteStep()

def compute_CTE_scale_factor(cte, parttemp):
    scaleFactor = getDoubleRef()
    sdk.SetStep("Compute CTE Scale Factor")
    sdk.SetDoubleArg("Material CTE (1/Deg F)", cte)
    sdk.SetDoubleArg("Initial Temperature (F)", parttemp)
    sdk.SetDoubleArg("Final Temperature (F)", 68.000000)
    sdk.ExecuteStep()
    sdk.GetDoubleArg("Scale Factor", scaleFactor)
    return scaleFactor

def set_instrument_scale_absolute(collection, instid, scaleFactor):
    sdk.SetStep("Set (absolute) Instrument Scale Factor (CAUTION!)")
    sdk.SetColInstIdArg("Instrument's ID", collection, instid)
    sdk.SetDoubleArg("Scale Factor", scaleFactor)
    sdk.ExecuteStep()
