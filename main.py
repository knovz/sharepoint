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


def folder_content_menu(
    sc: SharepointClient,
    title: str,
    drive_id: str,
    folder_id: str = None,
) -> dict:
    """
    Shows menu with folder content

    Arguments:
        content_list {list} -- folder contents

    Keyword Arguments:
        title {str} -- Optional title for the menu (default: {None})

    Returns:
        dict -- Selected item from the folder content
    """
    content_list = sc.list_folder_contents(drive_id, folder_id)

    items = []
    for item in content_list:
        item["displayName"] = item["name"]
        if "folder" in item:
            item["displayName"] += f" ({item["folder"]["childCount"]} items)"
        items.append(item)

    selected_item = cli_menu(
        items,
        title=title,
        prompt="Select one to continue",
        up_option=True,
        exit_option=True,
    )
    if "folder" in selected_item:
        selected_item = folder_content_menu(
            sc,
            f"{title} - {selected_item["name"]}",
            drive_id,
            selected_item["id"],
        )
        if selected_item["id"] == "UP":
            logger.info("User selected to move up one level")
            selected_item = folder_content_menu(
                sc,
                title,
                drive_id,
                folder_id,
            )

    return selected_item


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

    # SITES
    while True:
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

        # DRIVES
        while True:
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
                up_option=True,
                exit_option=True,
            )
            if selected_drive["id"] == "UP":
                logger.info("User selected to change site")
                break

            logger.info("User selected library: %s", selected_drive["name"])

            # FOLDERS
            selected_item = folder_content_menu(
                sc,
                f"{selected_site["displayName"]} - {selected_drive["name"]}",
                selected_drive["id"],
            )

            if selected_item["id"] != "UP":
                logger.info("User selected item: %s", selected_item["name"])
                break
            logger.info("User selected to change library")

        if selected_item["id"] != "UP":
            break

    print(f"User selected item: {selected_item["name"]}")


if __name__ == "__main__":
    main()
