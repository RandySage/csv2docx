from csv2docx import DocxConfig, utils
import unittest
from sys import stderr as err
import os
import csv
import inspect
import time

DEFAULT_INPUT_FILE = list(utils.locator('input.csv'))[0]

class TestDocxConfig(unittest.TestCase):

    def setUp(self):
        pass
    # end setUp

    def tearDown(self):
        pass

    def test_clean_changes_slash_r_ascii(self):
        test_text = 'aoeuu\naoeu\raoeu\r\naoeu'
        self.assertEqual(test_text.replace('\r', '\n'),
                         DocxConfig.clean(test_text))

    def test_clean_effect_on_unicode(self):
        with open(DEFAULT_INPUT_FILE) as input:
            test_text = input.read()
        self.assertEqual(test_text.replace('\r', '\n'),
                         DocxConfig.clean(test_text))


