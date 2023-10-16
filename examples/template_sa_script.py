import sys
import os
import logging


BASE_PATH = r"C:\Analyzer Data\Scripts\SA_Python_Lib"
LOG_FILE = os.path.join(BASE_PATH, "examples", "sa_debug_log.txt")


sys.path.append(os.path.join(BASE_PATH, "lib"))
import SAPyLib as sa


# #################################
# Your Methods and Classes go here
# #################################


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

    # Set the SA interaction mode
    sa.set_interaction_mode("Silent", "Never Halt", "Block Application Interaction")

    # ############################
    # Your script logic goes here
    # ############################

    # exit
    sys.exit(0)
