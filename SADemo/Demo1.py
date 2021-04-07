from SATools import *

try:
    # Data
    nominals = "Nominals"
    actuals = "Actuals"
    myGroup = "My Points"
    targetGroup = "Measured Points"
    singlePointProfile = "Single Pt. To SA"
    saInstrument = "Leica AT960/930"
    fitTolerance = 0.002

    myPoints = [Point3D(10, 10, 10),
                Point3D(10, 20, 10),
                Point3D(10, 10, 20),
                Point3D(5,  5,  5),
                Point3D(5,  10, 15),
                Point3D(15,  5, 20),
                ]

    # Procedure
    if not SAConnected:
        raise Exception('SA Not Connected')

    # cleanup job data
    objectNames = [
        "{}::{}::Point Group".format(nominals, myGroup),
        "{}::{}::Point Group".format(actuals, targetGroup)
        ]
    delete_objects(objectNames)

    # Re-setup collections 
    set_or_construct_default_collection(nominals)
    delete_collection(actuals)
    set_or_construct_default_collection(actuals)

    # Add Station 
    instid, _ = add_new_instrument(saInstrument)

    start_instrument(instid, False, True)

    # Do some work
    for i in range(len(myPoints)):
        pointName = "P" + str(i+1)

        construct_a_point(nominals, myGroup, pointName, 
                          myPoints[i].X,
                          myPoints[i].Y,
                          myPoints[i].Z)

        point_at_target(actuals, instid, nominals, myGroup, pointName)

        configure_and_measure(actuals, instid, actuals, targetGroup,
                              pointName, singlePointProfile,
                              True, True, 0)

    fitResult = best_fit_group_to_group(nominals, myGroup,
                                        actuals, targetGroup,
                                        True,
                                        fitTolerance, fitTolerance,
                                        False)

    stop_instrument(actuals, instid)

    if fitResult == mpresult.DONESUCCESS:
        print("Fit Good")
    else:
        print("Fit Failed {}".format(fitResult))

except Exception as err:
    print("Error: {}".format(str(err)))
