#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/7/3'

import os

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

chrome_profile_dir=r'C:/Users/rhj231223/AppData/Local/Google/Chrome/User Data/'
firefox_profile_dir=r'C:/Users/rhj231223/AppData/Roaming/Mozilla/Firefox/Profiles/egsctfvn.default/'

class SeleniumHelper(object):

    def __init__(self):
        self.browser=self.base_browser()

    def base_browser(self):
        options=webdriver.ChromeOptions()
        options.add_argument('--user-data-dir='+os.path.abspath(chrome_profile_dir))
        browser =webdriver.Chrome(chrome_options=options)
        # profile=webdriver.FirefoxProfile(firefox_profile_dir)
        # browser =webdriver.Firefox(profile)

        return browser

    def s_get(self,url,wait_css='.t_fsz img'):

        self.browser.get(url)
        wait=WebDriverWait(self.browser,10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,wait_css)))
        return self.browser.page_source


def b_get(browser,url,wait_css,max_load_time=15):

    browser.get(url)
    wait=WebDriverWait(browser,max_load_time)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,wait_css)))
    browser.execute_script("window.stop()")
    return browser

class SeleniumHelperNew(object):
    # def __init__(self):
    #     self.browser=self.base_browser()

    def base_browser(self):
        options=webdriver.ChromeOptions()
        options.add_argument('--user-data-dir='+os.path.abspath(chrome_profile_dir))
        browser =webdriver.Chrome(chrome_options=options)
        # profile=webdriver.FirefoxProfile(firefox_profile_dir)
        # browser =webdriver.Firefox(profile)

        return browser
