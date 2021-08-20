"""
## Download meeting recordings
Downloads webex or zoom recordings onto user provided paths.

## Usage:
> python download_webex_recordings.py --source <\path\to\source_file.csv> --driver <\path\to\chromedriver.exe> --destination <\path\to\downloads_dir>
- If destination is not provided, uses current directory for downloads.
- If driver is not provided, tries to search for chromedriver in current directory.
- Use --no-prompt/-n to avoid prompting user.

## Pre-requisites:
- Download relevant ChromeDriver from https://chromedriver.chromium.org/downloads place it under utilities folder.
- Install selenium using "pip install selenium".
- Use python version 3.6+ on Windows.

## More Info:
- If browser prompts from CAPTCHA (or) fails with invalid password:
    - When no-prompt mode is disabled: console will ask for user confirmation to proceed further after manually entering captcha.
    - When no-prompt mode is enabled: script execution will fail downloading for that particukar url and proceed further.
- If there is any exception while downloading a link, console shows that there is error downloading that particular link
and proceeds with other urls. At the end, all failed links will be displayed.
"""

import argparse
import csv
import sys
import time
import os

from time import sleep
from selenium import webdriver  # pip install selenium


class bcolors:
    PINK = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    ENDC = '\033[0m'


def print_warning(msg):
    print(f"{bcolors.YELLOW}{msg}{bcolors.ENDC}")


def print_failure(msg):
    print(f"{bcolors.RED}{msg}{bcolors.ENDC}")


def print_success(msg):
    print(f"{bcolors.GREEN}{msg}{bcolors.ENDC}")


class Downloader:
    webex_input_field_xpath = '//*[@id="meeting-password-form"]/div/div[1]/input'
    webex_ok_button_xpath = '//*[@id="meeting-password-form"]/button'
    webex_download_button_xpath = '//*[@class="icon-download recordingDownload"]'
    # webex_wrong_password_text = 'Invalid password. Try again.'

    zoom_input_field_xpath = '//*[@id="password"]'
    zoom_ok_button_xpath = '//*[@id="password_form"]/div/div/button'
    zoom_download_button_xpath = '//*[@class="download"]'
    # zoom_wrong_password_text = 'Wrong passcode'

    def __init__(self, urls_mapping, download_path, driver_path, no_prompt=False):
        self.driver_path = driver_path
        self.download_path = download_path
        self.driver = self.init_driver()
        self.urls_mapping = urls_mapping
        self.no_prompt = no_prompt
        self.failed_urls = []

    def init_driver(self):
        os.environ["webdriver.chrome.driver"] = self.driver_path
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_path,
            "download.prompt_for_donwload": False,
            'profile.default_content_setting_values.automatic_downloads': 1,
        }
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(self.driver_path, options=options)
        return driver

    def wait_until_chrome_finishes_download(self, timeout=600):
        print_warning(f'\nMake sure there are no old .crdownload files under {self.download_path}')
        print("Waiting for all downloads to finish. Waits until all .crdownload are removed from download directory")
        fileends = ".crdownload"
        seconds = 0
        should_we_wait = True
        while should_we_wait and seconds < timeout:
            time.sleep(1)
            should_we_wait = False
            files = os.listdir(self.download_path)
            for filename in files:
                if filename.endswith(fileends):
                    should_we_wait = True
            seconds += 1
        return seconds

    def get_website_type(self, url):
        if '.webex.com' in url:
            return 'webex'
        if 'zoom.us' in url:
            return 'zoom'
        raise Exception("Unknown website")

    def download_webex_link(self, url, password):
        try:
            self.driver.get(url)
            current_url = self.driver.current_url
            type = self.get_website_type(current_url)
            if type == 'webex':
                input_field_xpath = self.webex_input_field_xpath
                ok_button_xpath = self.webex_ok_button_xpath
                download_button_xpath = self.webex_download_button_xpath
            else:
                input_field_xpath = self.zoom_input_field_xpath
                ok_button_xpath = self.zoom_ok_button_xpath
                download_button_xpath = self.zoom_download_button_xpath

            print(f"\nAccessing {type} url: {current_url}")
            sleep(2)
            input_field = self.driver.find_element_by_xpath(input_field_xpath)
            input_field.send_keys(password)

            button_field = self.driver.find_element_by_xpath(ok_button_xpath)
            button_field.click()
            print('Wait for a while after clicking "OK"')
            sleep(5)

            download_elements = self.driver.find_elements_by_xpath(download_button_xpath)
            if len(download_elements) == 1:
                pass
            else:
                print_failure(
                    f'Unable to find download button yet. It is possible that {type} url is prompting for CAPTCHA or Wrong Password was entered.')
                if not self.no_prompt:
                    print_warning("Manually provide CATPCHA or valid credentials and go until download screen and press enter")
                    input('Press enter to proceed further: ')

            print('Initiating download')
            download_button = self.driver.find_element_by_xpath(download_button_xpath)
            download_button.click()
            print('Wait for download to get started')
            sleep(3)
        except Exception as e:
            self.failed_urls.append(url)
            print_failure(f'Error: {e}')
            print_failure('Unable to download')

    def download_all(self):
        print("Start processing...")
        try:
            for url, password in self.urls_mapping.items():
                self.download_webex_link(url, password)

            self.wait_until_chrome_finishes_download()
            if self.failed_urls:
                print_failure("\nUnable to download following urls:")
                for url in self.failed_urls:
                    print_failure(f"    - {url}")
        finally:
            print_success("Closing browser")
            self.driver.close()


def get_arg_parser():
    parser = argparse.ArgumentParser(usage=__doc__)
    parser.add_argument(
        '-s', '--source', help='Provide source file path with urls and passwords for recordings', required=True
    )
    parser.add_argument(
        '-d',
        '--destination',
        help='Provided directory for downloads',
    )
    parser.add_argument(
        '-c',
        '--driver',
        help='Provide webdriver path',
    )
    # Exit instead of displaying prompts. Useful for automation
    parser.add_argument('-n', '--no-prompt', help=argparse.SUPPRESS, action='store_true')
    return parser


def parse_csv(csv_filepath):
    urls_mapping = {}
    with open(csv_filepath) as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for index, row in enumerate(csv_reader):
            print(row)
            url = list(row.values())[0].strip()
            if url.startswith('#'):
                # probably a comment, so ignore
                continue
            password = list(row.values())[1].strip()
            urls_mapping[url] = password
    assert urls_mapping, "Unable to find required information from source file"
    return urls_mapping


def enable_vt100_support_on_windows():
    try:
        import ctypes
        kernel32 = ctypes.WinDLL('kernel32')
        hStdOut = kernel32.GetStdHandle(-11)
        mode = ctypes.c_ulong()
        kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
        mode.value |= 4
        kernel32.SetConsoleMode(hStdOut, mode)
    except Exception:
        # ignore exception
        pass


def main(args):
    enable_vt100_support_on_windows()
    parser = get_arg_parser()
    parsed_args = parser.parse_args(args)
    current_directory = os.path.split(os.path.realpath(__file__))[0]

    no_prompt = parsed_args.no_prompt
    csv_file = parsed_args.source

    download_dir = parsed_args.destination
    if not download_dir:
        download_dir = current_directory

    driver_path = parsed_args.driver
    if not driver_path:
        driver_path = os.path.join(current_directory, 'utilities', 'chromedriver.exe')
    assert os.path.exists(driver_path), f"{driver_path} do not exist. Please provide valid web driver path"

    urls_mapping = parse_csv(csv_file)
    downloader = Downloader(urls_mapping, download_dir, driver_path=driver_path, no_prompt=no_prompt)
    downloader.download_all()


DEBUG = False
if DEBUG:
    # Test Downloader
    urls_mapping = {
        '<url>': '<password>',
    }

    current_directory = os.path.split(os.path.realpath(__file__))[0]
    chromedriver_path = os.path.join(current_directory, 'utilities', 'chromedriver.exe')
    download_path = current_directory
    downloader = Downloader(urls_mapping, download_path, driver_path=chromedriver_path)
    downloader.download_all()

if __name__ == '__main__':
    main(sys.argv[1:])
