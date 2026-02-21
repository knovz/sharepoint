"""CLI MENU"""

import logging
import sys

from .clearscreen import clear_screen

logger = logging.getLogger(__name__)


def cli_menu(
    menu_options: list,
    title: str = None,
    prompt: str = "Select an option",
    up_option: bool = False,
    exit_option: bool = False,
    error_msg=None,
) -> dict:
    """
    Present menu options and get user input

    Arguments:
        menu_options {list} -- list of dictionary items.
                               Each item needs to have at least a "displayName" key.

    Keyword Arguments:
        title {str} -- optional title for the menu (default: {None})
        prompt {str} -- text to prompt for the user input (default: {"Select an option"})
        exit_option {bool} -- Include Exit option in the menu. (default: {False})

    Returns:
        dict -- _description_
    """

    while True:
        clear_screen()

        if title is not None:
            print(title)
        print("")
        print("")
        i = 0
        for item in menu_options:
            i += 1
            print(f"{i}. {item["displayName"]}")

        print("")
        if up_option:
            print("u. UP")
        if exit_option:
            print("x. EXIT")

        print("")
        if error_msg is not None:
            print(error_msg)
            error_msg = None
        print("")
        option = input(f"{prompt}\n")

        if option == "x" and exit_option:
            logger.debug("User selected to close the application")
            sys.exit(0)
        if option == "u" and up_option:
            logger.debug("User selected to move up one level")
            return {"id": "UP"}
        try:
            int_option = int(option)
            if int_option <= len(menu_options):
                logger.debug("User selected option: %s", int_option)
                return menu_options[int_option - 1]
            logger.debug("User selected option out of range: %s", int_option)
        except ValueError:
            logger.debug("Menu wrong user selection: %s", option)

        error_msg = " - Please select a valid option - "
