import string
import time
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
import os
import shutil
import glob
import ctypes
from ctypes import wintypes

try:
    import winreg
except ImportError:
    import _winreg as winreg

user32 = ctypes.WinDLL('user32', use_last_error=True)
gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)

FONTS_REG_PATH = r'Software\Microsoft\Windows NT\CurrentVersion\Fonts'

HWND_BROADCAST = 0xFFFF
SMTO_ABORTIFHUNG = 0x0002
WM_FONTCHANGE = 0x001D
GFRI_DESCRIPTION = 1
GFRI_ISTRUETYPE = 3

if not hasattr(wintypes, 'LPDWORD'):
    wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)

user32.SendMessageTimeoutW.restype = wintypes.LPVOID
user32.SendMessageTimeoutW.argtypes = (
    wintypes.HWND,   # hWnd
    wintypes.UINT,   # Msg
    wintypes.LPVOID, # wParam
    wintypes.LPVOID, # lParam
    wintypes.UINT,   # fuFlags
    wintypes.UINT,   # uTimeout
    wintypes.LPVOID  # lpdwResult
)

gdi32.AddFontResourceW.argtypes = (
    wintypes.LPCWSTR,) # lpszFilename

# http://www.undocprint.org/winspool/getfontresourceinfo
gdi32.GetFontResourceInfoW.argtypes = (
    wintypes.LPCWSTR, # lpszFilename
    wintypes.LPDWORD, # cbBuffer
    wintypes.LPVOID,  # lpBuffer
    wintypes.DWORD)   # dwQueryType


def get_links(page_url):
    headers = {"User-Agent": "Chrome",
               "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
               }
    response = requests.get(page_url, headers=headers)

    return BeautifulSoup(response.text, 'html.parser')


def create_main_urls():
    links = []
    for letter in list(string.ascii_lowercase):
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


def remove_zip_files(zip_files):
    # Ask user if they want to remove the .rar files
    remove_rars = input("Do you want to remove the .rar files? (y/n) ")
    if remove_rars.lower() == 'y':
        for zip_file in zip_files:
            os.remove(zip_file)


def unzip_fonts_from_downloads():
    downloads_dir = os.path.expanduser("~") + os.sep + "Downloads"
    temp_dir = downloads_dir + os.sep + "temp"
    shutil.rmtree(temp_dir, ignore_errors=True)
    os.makedirs(temp_dir, exist_ok=True)

    all_zip_files = glob.glob(downloads_dir + '/*.zip')
    zip_files = []
    for zip_file in all_zip_files:
        if "opti-" in zip_file:
            zip_files.append(zip_file)

    if not zip_files:
        print("No opti-*.zip files found in Downloads directory.")
        return

    # Loop through the opti-*.rar files
    for zip_file in zip_files:
        # Extract the contents of the .rar file to a temporary directory
        temp_dir = os.path.join(os.path.dirname(zip_file), 'temp')
        shutil.unpack_archive(zip_file, temp_dir)
    remove_zip_files(zip_files)


def install_font(src_path):
    # copy the font to the Windows Fonts folder
    dst_path = os.path.join(
        os.environ['SystemRoot'], 'Fonts', os.path.basename(src_path)
    )
    shutil.copy(src_path, dst_path)

    # load the font in the current session
    if not gdi32.AddFontResourceW(dst_path):
        os.remove(dst_path)
        raise WindowsError('AddFontResource failed to load "%s"' % src_path)

    # notify running programs
    user32.SendMessageTimeoutW(
        HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1000, None
    )

    # store the fontname/filename in the registry
    filename = os.path.basename(dst_path)
    fontname = os.path.splitext(filename)[0]

    # try to get the font's real name
    cb = wintypes.DWORD()
    if gdi32.GetFontResourceInfoW(
            filename, ctypes.byref(cb), None, GFRI_DESCRIPTION
    ):
        buf = (ctypes.c_wchar * cb.value)()
        if gdi32.GetFontResourceInfoW(
                filename, ctypes.byref(cb), buf, GFRI_DESCRIPTION
        ):
            fontname = buf.value

    is_truetype = wintypes.BOOL()
    cb.value = ctypes.sizeof(is_truetype)
    gdi32.GetFontResourceInfoW(
        filename, ctypes.byref(cb), ctypes.byref(is_truetype), GFRI_ISTRUETYPE
    )

    if is_truetype:
        fontname += ' (TrueType)'

    with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, FONTS_REG_PATH, 0, winreg.KEY_SET_VALUE
    ) as key:
        winreg.SetValueEx(key, fontname, 0, winreg.REG_SZ, filename)


def install_fonts_from_temp_folder():
    temp_dir = os.path.expanduser("~") + os.sep + "Downloads" + os.sep + "temp"
    # Ask user if they want to install the fonts
    install_fonts = input("Do you want to install all the fonts? (y/n) ")
    if install_fonts.lower() != 'y':
        return

    # Find any .ttf or .otf files in the temporary directory and install them
    font_files = glob.glob(temp_dir + '/*.ttf') + glob.glob(temp_dir + '/*.otf')
    for font_file in font_files:
        print("Installing:\t", font_file)
        install_font(font_file)

    # Delete the temporary directory
    delete_temp_folder = input("Do you want to delete the temp folder? (y/n) ")
    if delete_temp_folder.lower() != 'y':
        return
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


if __name__ == '__main__':
    main()
    unzip_fonts_from_downloads()
    install_fonts_from_temp_folder()

