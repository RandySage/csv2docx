from csv2docx import CsvParser, MySettings, JsonError, CrossRefError, utils, DocxConfig
import unittest
from sys import stderr as err
import os
import csv
import inspect
import time

THIS_FOLDER = os.path.abspath(os.path.dirname(__file__))

JSON_FILE = os.path.join(THIS_FOLDER, 'test_settings.json')
DEFAULT_INPUT_FILE = os.path.join(THIS_FOLDER, 'input.csv')
DEFAULT_OUTPUT_FILE = os.path.join(THIS_FOLDER, 'output.csv')


class TestMySettings(unittest.TestCase):

    def setUp(self):
        self.s = MySettings()
        self.s.read_json_file(JSON_FILE)
        self.s.INPUT_FILE = DEFAULT_INPUT_FILE
        self.s.OUTPUT_FILE = DEFAULT_OUTPUT_FILE

        self.parser = CsvParser(self.s)
    # end setUp

    def tearDown(self):
        pass

    def test_all_inds_includes_expected(self):
        """ Confirm mandatory index fields are in all_inds"""
        s = self.s
        self.assertTrue(s.id_ind in s.all_inds)
        self.assertTrue(s.heading_level_ind in s.all_inds)
        self.assertTrue(s.heading_num_ind in s.all_inds)
        self.assertTrue(s.heading_text_ind in s.all_inds)
        self.assertTrue(s.body_text_ind in s.all_inds)
    # test_all_inds_includes_expected


