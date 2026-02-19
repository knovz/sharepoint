"""
Minimal app to test Sharepoint access from python
"""

import os
import logging
import logging.config
import yaml
from dotenv import load_dotenv

from restclients import SharepointClient

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

    sc = SharepointClient(
        os.getenv("SHAREPOINT_CLIENT_ID"),
        os.getenv("SHAREPOINT_CLIENT_SECRET"),
        os.getenv("SHAREPOINT_TENANT_ID"),
    )

    logger.info(sc.app)

    sites = sc.list_sites()
    for site in sites:
        if "displayName" in site and not site["isPersonalSite"]:
            print(f"{site["displayName"]} - {site["id"]}")

    site = sc.get_site_by_name("captarvision.sharepoint.com", "ShareTest")
    print(f"{site["displayName"]} - {site["id"]}")

    # Works with both the long and the short version
    # site_id = "captarvision.sharepoint.com,02d9aeff-b1b8-40dd-ad98-cfb57a6953af,1850c9c8-fb1e-4b87-90ab-322c010209d8"
    # site_id = "02d9aeff-b1b8-40dd-ad98-cfb57a6953af"
    site_id = (
        "captarvision.sharepoint.com,"
        "02d9aeff-b1b8-40dd-ad98-cfb57a6953af,"
        "1850c9c8-fb1e-4b87-90ab-322c010209d8"
    )
    site = sc.get_site_by_id(site_id)
    if "error" in site:
        print_error(site["error"])
    else:
        print(f"{site["displayName"]} - {site["id"]}")


if __name__ == "__main__":
    main()
