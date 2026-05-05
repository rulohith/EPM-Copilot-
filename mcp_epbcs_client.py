import os
import base64
import requests
from dotenv import load_dotenv

load_dotenv()


class MCP_EPBCS_Client:
    def __init__(self):
        self.mcp_url = os.getenv("MCP_SERVER_URL")
        self.epbcs_url = os.getenv("EPBCS_BASE_URL")
        self.username = os.getenv("EPBCS_USERNAME")
        self.password = os.getenv("EPBCS_PASSWORD")

        if not all([self.mcp_url, self.epbcs_url, self.username, self.password]):
            raise ValueError("Missing required environment variables")

        # Create Basic Auth header (Base64 encoded)
        credentials = f"{self.username}:{self.password}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()

        self.headers = {
            "Authorization": f"Basic {encoded_credentials}",
            "Content-Type": "application/json"
        }

    def _call_mcp(self, method, endpoint, payload=None):
        """
        Generic MCP call wrapper
        """
        url = f"{self.mcp_url}{endpoint}"

        response = requests.request(
            method=method,
            url=url,
            headers=self.headers,
            json=payload
        )

        return self._handle_response(response)

    def get_application_details(self):
        """
        Example: Fetch EPBCS application details via MCP
        """
        payload = {
            "method": "GET",
            "url": self.epbcs_url
        }

        return self._call_mcp("POST", "/invoke", payload)

    def run_business_rule(self, rule_name):
        """
        Example: Execute business rule
        """
        payload = {
            "method": "POST",
            "url": f"{self.epbcs_url}/jobs",
            "body": {
                "jobType": "Rules",
                "jobName": rule_name
            }
        }

        return self._call_mcp("POST", "/invoke", payload)

    def _handle_response(self, response):
        if response.status_code in (200, 201):
            return response.json()
        else:
            raise Exception(
                f"Error {response.status_code}: {response.text}"
            )