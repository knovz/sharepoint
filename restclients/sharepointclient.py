"""Sharepoint Client"""

import json
import logging
import msal
import requests

logger = logging.getLogger(__name__)


class SharepointClient:
    """
    Class to hold a Client to connecto to a Sharepoint server
    """

    graph_api = "https://graph.microsoft.com/v1.0/{}"
    authority_base = "https://login.microsoftonline.com/{}"
    scope = ["https://graph.microsoft.com/.default"]

    def __init__(self, client_id, client_secret, tenant_id) -> None:
        self.token = None

        self.app = msal.ConfidentialClientApplication(
            client_id,
            client_credential=client_secret,
            authority=self.authority_base.format(tenant_id),
        )

        result = None
        result = self.app.acquire_token_silent(self.scope, account=None)
        if not result:
            logger.debug("No token on cache")
            result = self.app.acquire_token_for_client(self.scope)

        if "access_token" in result:
            logger.info("Acquired token with type '%s'", result["token_type"])
            self.token = result["access_token"]
        else:
            logger.error(
                "Error '%s' while requesting the token.\n    %s",
                result.get("error"),
                result.get("error_description"),
            )
            raise ConnectionError(
                f"Error '{result.get("error")}' while requesting the token."
                " Check the logs for more details."
            )

    def __get(self, graph_url: str, timeout: int = 5):
        response = requests.get(
            graph_url,
            headers={
                "Authorization": f"Bearer {self.token}",
            },
            timeout=timeout,
        )
        return response

    def __get_json(self, graph_url: str, timeout: int = 5):
        response = requests.get(
            graph_url,
            headers={
                "Authorization": f"Bearer {self.token}",
                "Accept": "application/json",
            },
            timeout=timeout,
        ).json()
        logger.debug(json.dumps(response, indent=2))
        return response

    def list_sites(self):
        """
        Retrieves all sites.
        """
        graph_url = "https://graph.microsoft.com/v1.0/sites/"
        #   id, displayName, name
        response = self.__get_json(graph_url)

        return response["value"]

    def get_site_by_name(self, host_name: str, site_name: str):
        """
        Retrieve a site by name and host name

        Arguments:
            host_name {str} -- name of the sharepoint site
            site_name {str} -- name of the site

        Returns:
            _type_ -- Object with site data
        """
        graph_url = (
            f"https://graph.microsoft.com/v1.0/sites/{host_name}:/sites/{site_name}"
        )
        response = self.__get_json(graph_url)
        return response

    def get_site_by_id(self, site_id: str):
        """
        Retrieve site by id

        Arguments:
            site_id {str} -- id of the sharepoint site

        Returns:
            _type_ -- Object with site data
        """
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        response = self.__get_json(graph_url)
        return response

    def get_site_drive(self, site_id: str) -> dict:
        """
        retrieve default document library for a site

        Arguments:
            site_id {str} -- site id

        Returns:
            dict -- Default library data
        """
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        reponse = self.__get_json(graph_url)
        return reponse

    def get_site_drives(self, site_id: str) -> list:
        """
        Retrieve a list of all document libraries for a site

        Arguments:
            site_id {str} -- site id

        Returns:
            list -- list of all document libraries for the site
        """
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        reponse = self.__get_json(graph_url)
        return reponse["value"]

    def get_drive_by_id(self, drive_id: str) -> dict:
        """
        Get drive information

        Arguments:
            drive_id {str} -- Drive id

        Returns:
            dict -- Drive information in dict form
        """
        # graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root"
        graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root"
        response = self.__get_json(graph_url)
        return response

    def list_drive_contents(self, drive_id: str) -> list:
        """
        List document library contents

        Arguments:
            drive_id {str} -- drive id

        Returns:
            list -- list of all driveItems in this drive root
        """
        graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        response = self.__get_json(graph_url)
        return response["value"]

    def list_folder_contents(self, drive_id: str, folder_id: str = None) -> list:
        """
        List folder contents

        Arguments:
            drive_id {str} -- document library id
            folder_id {str} -- folder id, None for root

        Returns:
            list -- list driveItems contained in this folder
        """
        if folder_id is not None:
            folder_path = f"/items/{folder_id}"
        else:
            folder_path = "/root"
        graph_url = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{folder_path}/children"
        )
        response = self.__get_json(graph_url)
        return response["value"]

    def download_file(self, file: dict) -> None:
        """
        download a file

        Arguments:
            file {dict} -- file, including id, name, and driveId parent reference
        """
        graph_url = f"https://graph.microsoft.com/v1.0/drives/{file["parentReference"]["driveId"]}/items/{file["id"]}/content"
        response = self.__get(graph_url)
        with open(f"out/{file["name"]}", "wb") as fd:
            for chunk in response.iter_content(chunk_size=128):
                fd.write(chunk)

    # Upload file
    # Replace
    # PUT /drives/{drive-id}/items/{item-id}/content
    # New file
    # PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content
    # with open("out/output.xlsx", "rb") as file:
    #     content = file.read()

    # response = requests.put(
    #     graph_url,
    #     headers=http_headers,
    #     timeout=5,
    #     data=content,
    # ).json()
    # logger.info(json.dumps(response, indent=2))
