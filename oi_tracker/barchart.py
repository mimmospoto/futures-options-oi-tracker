"""Fetch futures and futures-options data from barchart.com.

Uses a normal browser session (headless Chrome, driven by Selenium) and reads
the same JSON data the page itself loads. This is more reliable than parsing
the rendered HTML table.
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from urllib.parse import urlencode

from selenium import webdriver

BASE_URL = "https://www.barchart.com"
API_PATH = "/proxies/core-api/v1/quotes/get"
# Headless Chrome's default user agent (containing "HeadlessChrome") is refused
# by the site, so headless runs present the regular desktop Chrome string.
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/140.0 Safari/537.36"
)

# Runs inside the page: reads the XSRF cookie and calls the JSON API.
_FETCH_JS = """
const done = arguments[arguments.length - 1];
const match = document.cookie.match(/XSRF-TOKEN=([^;]+)/);
const token = match ? decodeURIComponent(match[1]) : '';
fetch(arguments[0], {
  headers: {'x-xsrf-token': token, 'accept': 'application/json'},
  credentials: 'include',
})
  .then(r => r.text())
  .then(done)
  .catch(e => done(JSON.stringify({error: String(e)})));
"""


class BarchartError(RuntimeError):
    pass


@dataclass(frozen=True)
class FuturesContract:
    symbol: str        # e.g. "KCZ26"
    last_price: float


class BarchartClient:
    """Context manager wrapping a headless Chrome session on barchart.com."""

    def __init__(self, headless: bool = True, page_load_wait: float = 6.0):
        self.headless = headless
        self.page_load_wait = page_load_wait
        self._driver = None

    def __enter__(self) -> "BarchartClient":
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument("--headless=new")
            options.add_argument(f"user-agent={USER_AGENT}")
        options.add_argument("--window-size=1400,900")
        # Selenium >= 4.6 downloads a matching chromedriver automatically.
        self._driver = webdriver.Chrome(options=options)
        self._driver.set_script_timeout(30)
        return self

    def __exit__(self, *exc) -> None:
        if self._driver is not None:
            self._driver.quit()
            self._driver = None

    def open(self, root: str) -> None:
        """Load a page for ``root`` so the session has valid cookies."""
        self._driver.get(f"{BASE_URL}/futures/quotes/{root}*0/options")
        time.sleep(self.page_load_wait)

    def _get(self, **params) -> list[dict]:
        url = f"{API_PATH}?{urlencode(params)}"
        text = self._driver.execute_async_script(_FETCH_JS, url)
        try:
            payload = json.loads(text)
        except (TypeError, json.JSONDecodeError) as exc:
            raise BarchartError(f"Non-JSON response for {url}: {text!r:.200}") from exc
        if "data" not in payload:
            raise BarchartError(f"Unexpected response for {url}: {payload}")
        return [row["raw"] for row in payload["data"]]

    def futures_contracts(self, root: str) -> list[FuturesContract]:
        """Listed futures contracts for ``root``, nearest expiry first."""
        rows = self._get(
            list="futures.contractInRoot", root=root,
            fields="symbol,lastPrice", raw=1,
        )
        return [
            FuturesContract(r["symbol"], float(r["lastPrice"]))
            for r in rows
            # "KCY00" style symbols are cash/nearby aggregates, not contracts.
            if not r["symbol"].endswith("00") and r.get("lastPrice") is not None
        ]

    def option_chain(self, contract: str) -> list[dict]:
        """All option strikes on a futures contract: strike, type, open interest."""
        return self._get(
            symbol=contract, list="futures.options",
            fields="strikePrice,optionType,openInterest",
            raw=1, limit=1000,
        )
