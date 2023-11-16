"""This Demo file shows the possibilities of using the Windows default Dialog Message Boxes."""

import sys
import os
import clr
import logging

BASE_PATH = r"C:\Analyzer Data\Scripts\SA_Python_Lib"
LOG_FILE = os.path.join(BASE_PATH, "examples", "sa_debug_log.txt")

sys.path.append(os.path.join(BASE_PATH, "lib"))
import SAPyLib as sa

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Windows.Forms as WinForms
from System.Drawing import Size, Point, Font


font = Font("Microsoft Sans Serif", 12.0)


class CTE:
    AluminumCTE_1DegF = 0.0000131
    SteelCTE_1DegF = 0.0000065
    CarbonFiberCTE_1DegF = 0


class MyDialog(WinForms.Form):
    """A simple hello world app that demonstrates the essentials of
    winforms programming and event-based programming in Python."""

    def __init__(self):
        self.Text = "Prompt"
        self.AutoScaleBaseSize = Size(5, 13)
        self.ClientSize = Size(400, 120)
        h = WinForms.SystemInformation.CaptionHeight
        self.MinimumSize = Size(400, (120 + h))

        # Create a label
        self.label_prompt = WinForms.Label()
        self.label_prompt.Text = "Question"
        self.label_prompt.Width = 150
        self.label_prompt.Location = Point(20, 20)

        # Create the text box
        self.textbox = WinForms.TextBox()
        self.textbox.Text = "Answer"
        self.textbox.TabIndex = 1
        self.textbox.Size = Size(180, 40)
        self.textbox.Location = Point(20, 48)

        # Create the button
        self.button = WinForms.Button()
        self.button.Text = "OK"
        self.button.TabIndex = 2
        self.button.Size = Size(80, 20)
        self.button.Location = Point(20, 84)

        # Register the event handler
        self.button.Click += self.button_Click

        # Add the controls to the form
        self.AcceptButton = self.button
        self.Controls.Add(self.button)
        self.Controls.Add(self.label_prompt)
        self.Controls.Add(self.textbox)

    def button_Click(self, sender, args):
        """Button click event handler"""
        log.debug("Click")
        self.DialogResult = WinForms.DialogResult.OK

    def get_user_input(self):
        return self.textbox.Text

    def set_question(self, question):
        self.label_prompt.Text = f"{question}:"


def user_query(prompt):
    """Ask the user a question, returns the answer."""
    log.debug("user_query")
    form = MyDialog()
    form.set_question(prompt)
    if form.ShowDialog() == WinForms.DialogResult.OK:
        user_input = form.get_user_input()
        form.Close()
        log.debug(f"User Input: {user_input}")
        return user_input


def test_run():
    """Test run the Class."""
    log.debug("test_run")
    answer = user_query("What is 5+5")
    if int(answer) == 10:
        print("OK")
    else:
        print("Failure")
    sys.exit(0)


if __name__ == "__main__":
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

    log.debug("start")

    sbGroup = "SB"
    simulate = True
    nominals = "Nominals"
    actuals = "Actuals"
    masterPlane1Points = "MasterPlane1"

    log.debug("Show the dialog")
    dialogResult = WinForms.MessageBox.Show("Start new job?", "Prompt", WinForms.MessageBoxButtons.OKCancel)
    if not dialogResult == WinForms.DialogResult.OK:
        raise OSError("User Aborted New Job")

    scalebarLength = 0.0
    while scalebarLength == 0.0 or scalebarLength < 80.0 or scalebarLength > 81.0:
        try:
            scalebarLength = float(user_query("Scalebar Length (in)"))
        except ValueError:
            scalebarLength = 0.0
    log.debug(f"scalebarLength: {scalebarLength}")

    partTemp = 0.0
    while partTemp == 0.0 or partTemp < 50.0 or partTemp > 120.0:
        try:
            partTemp = float(user_query("Part Temp (F)"))
        except ValueError:
            partTemp = 0.0
    log.debug(f"partTemp: {partTemp}")

    sa.rename_collection("A", nominals)

    if simulate:
        log.debug("Simulation!")
        sa.construct_a_point(nominals, "SB", "SB1", 50, 50, 0)
        sa.construct_a_point(nominals, "SB", "SB2", 50, 50 + scalebarLength, 0)

    sa.set_or_construct_default_collection(actuals)
    _, instid = sa.add_new_instrument("Leica AT960/930")
    scaleFactor = sa.compute_CTE_scale_factor(CTE.AluminumCTE_1DegF, partTemp)
    sa.set_instrument_scale_absolute(actuals, instid, scaleFactor)
    sa.start_instrument_interface("", instid, False, simulate)

    if simulate:
        log.debug("Simulation!")
        sa.point_at_target(actuals, instid, nominals, "SB", "SB1")

    log.debug("Show the dialog")
    WinForms.MessageBox.Show("Measure Scale Bar SB1", "Prompt", WinForms.MessageBoxButtons.OK)
    sa.configure_and_measure(actuals, instid, actuals, "SB BEG", "SB1", "Single Pt. To SA", True, True, 0)

    if simulate:
        log.debug("Simulation!")
        sa.point_at_target(actuals, instid, nominals, "SB", "SB2")

    log.debug("Show the dialog")
    WinForms.MessageBox.Show("Measure Scale Bar SB2", "Prompt", WinForms.MessageBoxButtons.OK)
    sa.configure_and_measure(actuals, instid, actuals, "SB BEG", "SB2", "Single Pt. To SA", True, True, 0)

    log.debug("Show the dialog")
    dialogResult = WinForms.MessageBox.Show("Measure Master Tool?", "Prompt", WinForms.MessageBoxButtons.OKCancel)
    if dialogResult == WinForms.DialogResult.OK:
        sa.configure_and_measure(actuals, instid, actuals, masterPlane1Points, "P1", "Stable Pt. To SA", True, True, 0)
    else:
        log.debug("Done")

    log.info("exit")
    sys.exit(0)
