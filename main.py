import string
import time
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
import os
import subprocess
import shutil
import glob


def get_links(page_url):
    headers = {"User-Agent": "Chrome",
               "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
               }
    response = requests.get(page_url, headers=headers)

    return BeautifulSoup(response.text, 'html.parser')


def create_main_urls():
    links = []
    my_Test_List = ["n", "m", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"]

    #for letter in list(string.ascii_lowercase):
    for letter in my_Test_List:
        url = "http://abfonts.freehostia.com/opti/fonts-" + letter + "/index.htm"
        links.append(url)
    return links


def read_font_urls(myList, index):
    fonts_download_links = []
    for linky in get_links(myList[index]).find_all('a'):
        href = linky.get('href')
        fonts_download_links.append(href)
    return fonts_download_links


def show_letter_fonts(links, index):
    things = []
    for font in read_font_urls(links, index):
        if font is None:
            continue
        if "download.htm" not in font:
            continue
        else:
            first_letter = ""
            if "#" in font:
                first_letter = font.split("#")[1]
            elif "?" in font:
                first_letter = font.split("?")[1]

            if first_letter == "":
                continue
            font = font[2:]
            url = "http://abfonts.freehostia.com/opti/fonts-" + first_letter[0].lower() + "/index.htm" + font
            things.append(url)
    return things


def click_link(driver, link_url):
    link = driver.find_element(By.XPATH, f'//a[@href="{link_url}"]')
    link.click()
    driver.execute_script("javascript:void(0);")
    time.sleep(1)  # Wait for the download button to appear
    try:
        download_button = driver.find_element(By.XPATH, '//a[contains(text(), "Download")]')
        download_button.click()
        time.sleep(2)  # Wait for the download to complete
        driver.switch_to.window(driver.window_handles[-1])  # switch to the popup window
        popup_close_button = driver.find_element(By.XPATH, '//span/a/img[@alt="Close"]')
        popup_close_button.click()
        driver.switch_to.window(driver.window_handles[0])  # switch back to the main window
    except NoSuchElementException:
        pass


def install_fonts_from_downloads():
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    # Search for .rar files in the downloads directory
    rar_files = glob.glob(downloads_dir + '/*.rar')

    # Loop through the .rar files
    for rar_file in rar_files:
        # Extract the contents of the .rar file to a temporary directory
        temp_dir = os.path.join(os.path.dirname(rar_file), 'temp')
        shutil.rmtree(temp_dir, ignore_errors=True)
        os.makedirs(temp_dir, exist_ok=True)
        subprocess.call(['unrar', 'x', '-o+', rar_file, temp_dir])

        # Find any .ttf or .otf files in the temporary directory and install them
        font_files = glob.glob(temp_dir + '/*.ttf') + glob.glob(temp_dir + '/*.otf')
        for font_file in font_files:
            subprocess.call(['powershell.exe', '-Command', 'Install-Font', font_file])

        # Delete the temporary directory
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    links = create_main_urls()
    for i in range(0, len(links) - 1):
        driver = webdriver.Chrome()
        driver.get(links[i])

        font_links = show_letter_fonts(links, i)
        for font_link in font_links:
            current_link = "../" + font_link.split("/")[6]
            click_link(driver, current_link)
        driver.quit()

    # install_fonts_from_downloads()


if __name__ == '__main__':
    main()
    # install_fonts_from_downloads()

