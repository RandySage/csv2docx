from csv2docx import CsvParser, MySettings, JsonError, CrossRefError, utils, DocxConfig
import unittest
from sys import stderr as err
import os
import csv
import inspect
import time

THIS_FOLDER = os.path.abspath(os.path.dirname(__file__))

JSON_FILE = os.path.join(THIS_FOLDER, 'test_settings.json')
DEFAULT_INPUT_FILE = os.path.join(THIS_FOLDER,'input.csv')
DEFAULT_OUTPUT_FILE = os.path.join(THIS_FOLDER,'output.csv')


class TestEndEnd(unittest.TestCase):
  
    def setUp(self):
        self.s = MySettings()
        self.s.read_json_file(JSON_FILE)
        self.s.INPUT_FILE = DEFAULT_INPUT_FILE
        self.s.OUTPUT_FILE = DEFAULT_OUTPUT_FILE
        
        self.parser = CsvParser(self.s)
    # end setUp

    def tearDown(self):
        pass
    
    def test_produces_output(self):
        out_docx = DocxConfig(self.s)
    
        self.parser.write_docx(out_docx)

        out_docx.save(self.s.OUTPUT_FILE)
        
        #At least confirm the file was modified
        num_seconds = 3
        self.assertTrue(time.time() - os.stat(self.s.OUTPUT_FILE).st_mtime <
                        num_seconds)




