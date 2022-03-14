import sys
import logging

sys.path.append("C:/Analyzer Data/Scripts/SAPython/lib")
import SAPyLib as sa


LOG_FILE = "C:/Analyzer Data/Scripts/sa_debug_log.txt"
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


# #################################
# Your Methods and Classes go here
# #################################


if __name__ == "__main__":
    # Set the SA interaction mode
    sa.set_interaction_mode("Silent", "Never Halt", "Block Application Interaction")

    # ############################
    # Your script logic goes here
    # ############################

    # exit
    sys.exit(0)
