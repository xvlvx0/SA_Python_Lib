"""
This example creates a Plane Relationship as a base task.
Additionally it can create a callout.
"""
import sys
import logging

sys.path.append("C:\Analyzer Data\Scripts\SAPython\lib")
import SAPyLib as sa
import tasks

LOG_FILE = "C:/Analyzer Data/Scripts/SAPython/sa_debug_log.txt"
logging.basicConfig(
    level=logging.DEBUG,
    filename=LOG_FILE,
    format="%(asctime)-12s - %(name)-8s - %(levelname)s - %(message)s",
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(name)-8s - %(levelname)-8s - %(message)s"))
logging.getLogger("").addHandler(console)
log = logging.getLogger(__name__)


if __name__ == "__main__":
    nomCollection = "Nominals"  # Nominals Collection name
    nomGroup = "NomPoints"  # Nominal Point Group
    nomPlane = "NomPlane"  # Nominal Plane name
    collection = "MeasuredData"  # destination collection
    measuredgroup = "MyPlane_measured"  # The group containing the measured points
    relname = "MyPlane"  # relationshipname

    # create the nominals collection
    sa.set_or_construct_default_collection(nomCollection)

    # create nominal points
    # point data in mm
    myPoints = [
        sa.NamedPoint3D("p1", xyz=[100, 100, 100]),
        sa.NamedPoint3D("p2", xyz=[100, 200, 100]),
        sa.NamedPoint3D("p3", xyz=[100, 100, 200]),
        sa.NamedPoint3D("p4", xyz=[50, 50, 50]),
        sa.NamedPoint3D("p5", xyz=[50, 100, 150]),
        sa.NamedPoint3D("p6", xyz=[150, 50, 200]),
    ]
    for point in myPoints:
        sa.construct_a_point_in_working_coordinates(nomCollection, nomGroup, point.name, point.X, point.Y, point.Z)

    # Construct plane
    sa.construct_plane(nomCollection, nomPlane)

    # Add Tracker
    sa.set_or_construct_default_collection(collection)
    instrument_Col, instrument_ID = sa.add_new_instrument("Leica AT960/930")
    sa.start_instrument_interface(instrument_Col, instrument_ID, False, True)

    # select the points to be measured
    pointlist = sa.make_a_point_name_ref_list_from_a_group(nomCollection, nomGroup)  # returns a points object

    # measure the nominal points
    measurement = tasks.measurePoints(measuredcollection=collection, measuredgroup=measuredgroup, points=pointlist)
    measurement.measure()

    sa.set_or_construct_default_collection(collection)
    # Make a plane relationship.
    # Inputs are:
    # - The destination collection name for the relationship
    # - The relationship name
    # - The measured point group name (can be an empty group)
    # - The nominal geometry [optional]
    myPlane = tasks.makeRelationshipPlane(
        relcollection=collection,
        relname=relname,
        measuredcollection=collection,
        measuredgroup=measuredgroup,
        nomcollection=nomCollection,
        nomobj=nomPlane,
    )
    myPlane.make()
    # myPlane.addCallOut()

    sys.exit(0)
