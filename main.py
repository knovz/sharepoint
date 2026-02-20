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

logger = logging.getLogger(__name__)
load_dotenv()


def clear_screen() -> None:
    """
    Send clear screen cli command depending on Operating system.
    https://stackoverflow.com/a/684344/4220807
    """
    os.system("cls" if os.name == "nt" else "clear")


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
    clear_screen()

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

    sites = sc.list_sites()
    for site in sites:
        if "displayName" in site and not site["isPersonalSite"]:
            print(f"{site["displayName"]} - {site["id"]}")

    site = sc.get_site_by_name("captarvision.sharepoint.com", "ShareTest")
    print(f"{site["displayName"]} - {site["id"]}")


if __name__ == "__main__":
    main()
