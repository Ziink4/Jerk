import os
import re
from typing import List

from bs4 import BeautifulSoup
import requests
from logzero import logger

SITEMAP_URL = "https://bikez.com/sitemap/motorcycle-specs.xml"
URL_LIST_FILENAME = "url.list"


def parse_sitemap(sitemap: str) -> List[str]:
    logger.info(f"Retrieving sitemap URLs from {sitemap}")

    if os.path.isfile(URL_LIST_FILENAME):
        # Load URL list from cache
        with open(URL_LIST_FILENAME, "r") as f:
            url_list = f.readlines()
            logger.info(f"Retrieved {len(url_list)} URLs from cache")

    else:
        # Get URL list and save it to the cache
        r = requests.get(sitemap)
        soup = BeautifulSoup(r.text, 'xml')
        url_list = [url.loc.string for url in soup.urlset.find_all("url")]
        with open(URL_LIST_FILENAME, "w+") as f:
            f.write('\n'.join(url_list))
            logger.info(f"Retrieved and cached {len(url_list)} URLs")

    return url_list


def retrieve_page(url: str):
    match_name = R".+/(.+)\..+"
    name = re.match(match_name, url)[1]
    # logger.info(f"Retrieving page {name}")


def retrieve_data(url_list: List[str]):
    for url in url_list:
        retrieve_page(url)


if __name__ == '__main__':
    urls = parse_sitemap(SITEMAP_URL)
    retrieve_data(urls)
