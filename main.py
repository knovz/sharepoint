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


if __name__ == "__main__":
    main()
