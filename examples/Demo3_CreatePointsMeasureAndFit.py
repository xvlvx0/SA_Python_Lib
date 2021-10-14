"""
This example creates a Plane Relationship as a base task.
Additionally it can create a callout.
"""
import sys

sys.path.append("C:\Analyzer Data\Scripts\SAPython\lib")  # adding the library to the path variable
import tasks
import SAPyLib as sa


if __name__ == "__main__":
    nomCollection = "Nominals"  # Nominals Collection name
    nomGroup = "NomPoints"  # Nominal Point Group
    collection = "MeasuredData"  # destination collection
    measuredgroup = "MyPlane_measured"  # The group containing the measured points
    relname = "MyPlane"  # relationshipname

    # create the nominals collection
    sa.set_or_construct_default_collection(nomCollection)

    # create nominals collection
    # point data in mm
    myPoints = [
        sa.NamedPoint3D("p1", 100, 100, 100),
        sa.NamedPoint3D("p2", 100, 200, 100),
        sa.NamedPoint3D("p3", 100, 100, 200),
        sa.NamedPoint3D("p4", 50, 50, 50),
        sa.NamedPoint3D("p5", 50, 100, 150),
        sa.NamedPoint3D("p6", 150, 50, 200),
    ]
    for point in myPoints:
        sa.construct_a_point(nomCollection, nomGroup, point.name, point.X, point.Y, point.Z)

    # Add Tracker
    sa.set_or_construct_default_collection(collection)
    instrument = sa.add_new_instrument("Leica AT960/930")
    sa.start_instrument(instrument[1], instrument[2], False, True)

    # Create a points list
    pointlist = []
    for point in myPoints:
        pointlist.append([nomCollection, nomGroup, point.name])

    print(f"\nPointlist: {pointlist}")

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
        nomcollection="Nominals",
        nomobj="NomPlane",
    )
    myPlane.make()
    # myPlane.addCallOut()

    sys.exit(0)
