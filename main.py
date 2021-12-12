import os
import re
from typing import List, Dict, Tuple, Optional

import logzero
from bs4 import BeautifulSoup
import requests
from logzero import logger
from tqdm import tqdm
from multiprocessing import Pool

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SITEMAP_URL = "https://bikez.com/sitemap/motorcycle-specs.xml"
URL_LIST_FILENAME = "url.list"

N_PROCESS = 25


def parse_sitemap(sitemap: str) -> List[str]:
    logger.info(f"Retrieving sitemap URLs...")

    if os.path.isfile(URL_LIST_FILENAME):
        # Load URL list from cache
        with open(URL_LIST_FILENAME, "r") as f:
            url_list = f.read().splitlines()
            logger.info(f"Retrieved {len(url_list)} URLs from cache.")
    else:
        # Get URL list and save it to the cache
        r = requests.get(sitemap)
        soup = BeautifulSoup(r.text, "xml")
        url_list = [url.loc.string for url in soup.urlset.find_all("url")]
        with open(URL_LIST_FILENAME, "w+") as f:
            f.write('\n'.join(url_list))
            logger.info(f"Retrieved and cached {len(url_list)} URLs.")

    return url_list


def retrieve_entry(soup: BeautifulSoup, key_text: str):
    key = soup.find(string=key_text)
    if key is None:
        return None

    if key.parent.parent.next_sibling is None:
        # Special case for table entries that are also links
        return key.parent.parent.parent.next_sibling.text

    return key.parent.parent.next_sibling.text


def parse_power(soup: BeautifulSoup) -> Tuple[Optional[float], Optional[float]]:
    power_entry = retrieve_entry(soup, "Power:")
    if power_entry is None:
        return None, None

    regex = R"(.+) HP \((.+)  kW\)\).*"
    match = re.match(regex, power_entry)
    return float(match[1]), float(match[2])


def parse_weight(soup: BeautifulSoup, key_text: str) -> Tuple[Optional[float], Optional[float]]:
    weight_entry = retrieve_entry(soup, key_text)
    if weight_entry is None:
        return None, None

    regex = R"(.+) kg \((.+) pounds\).*"
    match = re.match(regex, weight_entry)
    return float(match[1]), float(match[2])


def parse_power_weight_ratio(soup: BeautifulSoup) -> Optional[float]:
    p_w_r_entry = retrieve_entry(soup, "Power/weight ratio:")
    if p_w_r_entry is None:
        return None

    regex = R"(.+) HP/kg"
    match = re.match(regex, p_w_r_entry)
    return float(match[1])


def retrieve_page(url: str):
    r = requests.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    power_hp, power_kw = parse_power(soup)
    wet_weight_kg, wet_weight_lb = parse_weight(soup, "Weight incl. oil, gas, etc:")
    dry_weight_kg, dry_weight_lb = parse_weight(soup, "Dry weight:")

    entries = {"model": retrieve_entry(soup, "Model:"),
               "year": int(retrieve_entry(soup, "Year:")),
               "power_hp": power_hp,
               "power_kw": power_kw,
               "torque": retrieve_entry(soup, "Torque"),
               "displacement": retrieve_entry(soup, "Displacement"),
               "wet_weight_kg": wet_weight_kg,
               "wet_weight_lb": wet_weight_lb,
               "dry_weight_kg": dry_weight_kg,
               "dry_weight_lb": dry_weight_lb,
               "power_weight_ratio_hp_kg": parse_power_weight_ratio(soup),
               "url": url}
    return entries


def retrieve_data(url_list: List[str]) -> List[Dict[str, str]]:
    logger.debug("Starting the pool.")
    with Pool(N_PROCESS) as pool:
        results = []
        for result in tqdm(pool.imap_unordered(retrieve_page, url_list), total=len(url_list)):
            results.append(result)
        return results


def export_data(data: List[Dict[str, str]]) -> None:
    logger.info("Exporting data as Excel spreadsheet")
    wb = Workbook()
    ws = wb.active

    ws.append(list(data[0].keys()))
    for row in data:
        ws.append(list(row.values()))

    tab = Table(displayName="Data", ref="A1:" + get_column_letter(ws.max_column) + str(ws.max_row))
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9",
                                        showFirstColumn=False,
                                        showLastColumn=False,
                                        showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tab)
    wb.save("export.xlsx")


if __name__ == "__main__":
    urls = parse_sitemap(SITEMAP_URL)
    data = retrieve_data(urls[:100])
    export_data(data)
