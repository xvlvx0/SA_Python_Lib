"""
This Demo1 file demonstrates the SAPyLib functionality.
A very basic use case of several functions:
- declaring variables
- creation of points, groups and collections
- deletion of points, groups and collections
- adding and connecting to an instrument
- pointing at and measuring the nominal points
- basic analysis of the measured data

All 3 of the 'Demo1.*' files work together.
The Demo1.xit64 provides the basic setup.
The Demo1.mp files is connected to a button in the scripts section of the Demo1.xit64 file and calls the Demo1.py file.
This Demo1.py file executes the scripted logic.
"""
import sys

sys.path.append("C:\Analyzer Data\Scripts\SAPython\lib")  # adding the library to the path variable
import SAPyLib as sa  # importing the library


if __name__ == "__main__":
    # Declaring variables
    nomCol = "Nominals"
    nomGroup = "NomPoints"
    actualCol = "Actuals"
    actualGroup = "Measured Points"
    singlePointProfile = "Single Pt. To SA"
    saInstrument = "Leica AT960/930"
    fitTolerance = 0.2  # mm

    # create a point list
    # (values in mm)
    myPoints = [
        sa.Point3D(100, 100, 100),
        sa.Point3D(100, 200, 100),
        sa.Point3D(100, 100, 200),
        sa.Point3D(50, 50, 50),
        sa.Point3D(50, 100, 150),
        sa.Point3D(150, 50, 200),
    ]

    # Below are three different ways of deleting objects:
    # - by individual points
    # - by point groups
    # - by collections
    #
    # delete points from collection and pointgroup
    sa.delete_points_wildcard_selection(nomCol, nomGroup, "Point Group", "p*")
    #
    # delete point groups
    sa.delete_objects(actualCol, actualGroup, "Point Group")
    #
    # delete collections
    sa.delete_collection(nomCol)
    sa.delete_collection(actualCol)

    # Creating collections
    sa.set_or_construct_default_collection(nomCol)
    sa.set_or_construct_default_collection(actualCol)

    # Add an instrument and connect
    instCol, instId = sa.add_new_instrument(
        saInstrument
    )  # two variables are returned, the collection and object names of the tracker
    sa.start_instrument_interface(instCol, instId, False, True)  # run the tracker connection method

    # Create and measure the nominal points
    for i in range(len(myPoints)):
        # make a point
        pointName = f"p{i+1}"
        sa.construct_a_point(nomCol, nomGroup, pointName, myPoints[i].X, myPoints[i].Y, myPoints[i].Z)
        # point at
        sa.point_at_target(instCol, instId, nomCol, nomGroup, pointName)

        # measure point
        sa.configure_and_measure(instCol, instId, actualCol, actualGroup, pointName, singlePointProfile, True, True, 0)

    # fit the measured points
    fitResult = sa.best_fit_transformation_group_to_group(
        nomCol, nomGroup, actualCol, actualGroup, True, fitTolerance, fitTolerance, False
    )
    if fitResult:
        print("Fit Good")
    else:
        print(f"Fit Failed {fitResult}")

    # Disconnect the instrument
    sa.stop_instrument_interface(instCol, instId)

    # Exit the script
    print("Demo1 finished!")
    sys.exit(0)
