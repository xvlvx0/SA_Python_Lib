import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox, MessageBoxButtons, DialogResult
from Forms import user_query

try:
    sbGroup = "SB"
    simulate = True
    nominals = "Nominals"
    actuals = "Actuals"
    masterPlane1Points = "MasterPlane1"

    dialogResult = MessageBox.Show("Start new job?", "Prompt", MessageBoxButtons.OKCancel)
    if not dialogResult == DialogResult.OK:
        raise Exception("User Aborted New Job")

    from SATools import *

    if not SAConnected:
        raise Exception('SA Not Connected')

    scalebarLength = 0
    while scalebarLength == 0 or scalebarLength < 80 or scalebarLength > 81:
        try:
            scalebarLength = float(user_query("Scalebar Length (in)"))
        except:
            scalebarLength = 0

    partTemp = 0
    while partTemp == 0 or partTemp < 50 or partTemp > 120:
        try:
            partTemp = float(user_query("Part Temp (F)"))
        except:
            partTemp = 0

    rename_collection("A", nominals)

    if simulate:
        construct_a_point(nominals, "SB", "SB1", 50, 50, 0)
        construct_a_point(nominals, "SB", "SB2", 50, 50 + scalebarLength, 0)

    set_or_construct_default_collection(actuals)
    instid, _ = add_new_instrument("Leica AT960/930")
    scaleFactor = compute_CTE_scale_factor(CTE.AluminumCTE_1DegF, partTemp)
    set_instrument_scale_absolute(actuals, instid, scaleFactor)
    start_instrument(instid, False, simulate)

    if simulate:
        point_at_target(actuals, instid, nominals, "SB", "SB1")

    MessageBox.Show("Measure Scale Bar SB1", "Prompt", MessageBoxButtons.OK)
    configure_and_measure(actuals, instid, actuals, "SB BEG", "SB1",
                          "Single Pt. To SA", True, True, 0)

    if simulate:
        point_at_target(actuals, instid, nominals, "SB", "SB2")

    MessageBox.Show("Measure Scale Bar SB2", "Prompt", MessageBoxButtons.OK)
    configure_and_measure(actuals, instid, actuals, "SB BEG", "SB2",
                          "Single Pt. To SA", True, True, 0)

    dialogResult = MessageBox.Show("Measure Master Tool?", "Prompt", MessageBoxButtons.OKCancel)
    if dialogResult == DialogResult.OK:
        configure_and_measure(actuals, instid, actuals, masterPlane1Points, "P1",
                              "Stable Pt. To SA", True, True, 0)
    else:
        raise Exception("Done")

except Exception, e:
    print "Error: {}".format(str(e))




