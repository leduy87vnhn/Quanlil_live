from urllib.parse import urlencode
import requests

class VmixClient:
    def __init__(self, host: str = "127.0.0.1", port: int = 8088, timeout: int = 5):
        self.host = host
        self.port = port
        self.base = f"http://{self.host}:{self.port}/API/"
        self.timeout = timeout

    def call_api(self, params: dict) -> str:
        url = self.base + "?" + urlencode(params)
        r = requests.get(url, timeout=self.timeout)
        r.raise_for_status()
        return r.text

    def set_text(self, selected_name: str, value: str, input_index: int = None) -> str:
        params = {"Function": "SetText", "SelectedName": selected_name, "Value": value}
        if input_index is not None:
            params["Input"] = input_index
        return self.call_api(params)

    def get_state(self) -> str:
        r = requests.get(self.base, timeout=self.timeout)
        r.raise_for_status()
        return r.text
