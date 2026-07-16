from playwright.sync_api import sync_playwright
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

URL = (
    "https://meteo.arso.gov.si/uploads/probase/www/observ/"
    "surface/text/sl/observationAms_KRANJ_history.html"
)

MAPA = Path("screenshots")


def naredi_posnetek():
    MAPA.mkdir(parents=True, exist_ok=True)

    zdaj = datetime.now(ZoneInfo("Europe/Ljubljana"))
    datum = zdaj.strftime("%Y-%m-%d")
    ura = zdaj.strftime("%H-%M-%S")

    izhod = MAPA / f"KRANJ_{datum}_{ura}.png"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        page = browser.new_page(
            viewport={
                "width": 1400,
                "height": 2200,
            }
        )

        page.goto(
            URL,
            wait_until="networkidle",
            timeout=60000,
        )

        page.screenshot(
            path=str(izhod),
            full_page=True,
        )

        browser.close()

    print(f"Shranjeno: {izhod}")


if __name__ == "__main__":
    naredi_posnetek()
