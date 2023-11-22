import time
import math
import sys
from PyQt5.QtWidgets import *
from selenium.webdriver import ActionChains
from PyQt5 import uic
from selenium.webdriver.common.by import By
from selenium import webdriver
from twocaptcha import TwoCaptcha

from PyQt5.QtCore import *

class EverLogin(QThread):
    def __init__(self,driver):
        super().__init__()
        self.running = True
        self.driver =driver
        self.Can = True

    def run(self):
        try:
            self.driver.get('https://www.everugg.com/oms/login')
        except Exception as e:
            print(e)

    def resume(self):
        self.running = True

    def pause(self):
        self.running = False


class OzLogin(QThread):
    def __init__(self, driver):
        super().__init__()
        self.running = True
        self.driver =driver
        self.Can = True

    def run(self):
        try:
            self.driver.get('http://df.ozwearugg.com.au/login')

        except Exception as e:
            print(e)

    def resume(self):
        self.running = True

    def pause(self):
        self.running = False
