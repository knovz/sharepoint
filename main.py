"""
Minimal app to test Sharepoint access from python
"""

import os
import sys
import logging
import logging.config
import yaml
from dotenv import load_dotenv

from restclients import SharepointClient
from helpers import cli_menu

logger = logging.getLogger(__name__)
load_dotenv()


def print_error(error):
    """
    Prints and logs request error

    Arguments:
        error {_type_} -- error object with code and message
    """
    error_string = "{} - {}".format(error["code"], error["message"])
    logger.error(error_string)
    print(error_string)


def main() -> None:
    """
    Main entry point
    """
    with open("logging_config.yaml", "r", encoding="utf-8") as file:
        log_conf = yaml.safe_load(file)
    logging.config.dictConfig(log_conf)

    logger.info("Running")

    try:
        sc = SharepointClient(
            os.getenv("SHAREPOINT_CLIENT_ID"),
            os.getenv("SHAREPOINT_CLIENT_SECRET"),
            os.getenv("SHAREPOINT_TENANT_ID"),
        )
    except ConnectionError as ce:
        sys.exit(f"Could not connect to Sharepoint. {ce}")

    logger.info(sc.app)

    print("Connected to Sharepoint")

    sites = []
    site_list = sc.list_sites()

    for site in site_list:
        if "displayName" in site and not site["isPersonalSite"]:
            sites.append(site)

    selected_site = cli_menu(
        sites,
        title="Sharepoint sites",
        prompt="Please select one to continue",
        exit_option=True,
    )

    logger.info("User selected site: %s", selected_site["displayName"])

    drives = []
    drives_list = sc.get_site_drives(selected_site["id"])

    for drive in drives_list:
        if "name" in drive and not "displayName" in drive:
            drive["displayName"] = drive["name"]
        drives.append(drive)

    selected_drive = cli_menu(
        drives,
        title=f"{selected_site["displayName"]} libraries",
        prompt="Please select one to continue",
        exit_option=True,
    )

    logger.info("User selected library: %s", selected_drive["name"])
    print("User selected library: %s", selected_drive["name"])


if __name__ == "__main__":
    main()
