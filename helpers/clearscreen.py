"""clear_screen"""

import os


def clear_screen() -> None:
    """
    Send clear screen cli command depending on Operating system.
    https://stackoverflow.com/a/684344/4220807
    """
    os.system("cls" if os.name == "nt" else "clear")
