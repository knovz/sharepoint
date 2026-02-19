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

    graph_api = "https://graph.microsoft.com/v1.0/"
    authority_base = "https://login.microsoftonline.com/{}"
    scope = ["https://graph.microsoft.com/.default"]

    def __init__(self, client_id, client_secret, tenant_id) -> None:
        self.http_headers = None

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
            self.http_headers = {
                "Authorization": f'Bearer {result["access_token"]}',
                "Accept": "application/json",
            }
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

    def list_sites(self):
        """
        Retrieves all sites.
        """
        graph_url = "https://graph.microsoft.com/v1.0/sites/"
        #   id, displayName, name

        response = requests.get(
            graph_url,
            headers=self.http_headers,
            timeout=5,
        ).json()
        logger.debug(json.dumps(response, indent=2))
        return response["value"]
