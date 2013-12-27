#! /usr/bin/env python

"""
This file makes a .docx (Word 2007) file from scratch, showing off most of the
features of python-docx.

If you need to make documents from scratch, you can use this file as a basis
for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

DEFAULT_JSON = 'settings.json'

from docx import *
import re
import csv
import sys
import json
import inspect

class MySettings():

    def get_json(self, json_filename):
        try:
            with open(json_filename) as json_file:
                json_dict = json.loads(json_file.read())
                return json_dict
        except ValueError:
            sys.exit("Parse error for json file: %s\nExiting..." % json_filename)
        except IOError:
            sys.exit("Exiting from %s because failed to open json file: %s\nExiting..." % 
                     (inspect.stack()[0][3],json_filename))
    # get_json
    
    def read_json(self):
        try:
            settings = self.get_json(DEFAULT_JSON)
            # try:
            #     s.l_delim = settings.reference.start
            #     s.r_delim = settings.reference.end
            # except:
            #     sys.exit("Exiting with issue in get_settings\nExiting...")
            if ( len(settings['l_delim']) > 1 or
                 len(settings['r_delim']) > 1 ):
                sys.exit(
                    ("l_delim ('%s') and r_delim ('%s') need to be single characters\nExiting..." % 
                     (settings['l_delim'], settings['r_delim'])))
            self.l_delim = settings['l_delim'] 
            self.r_delim = settings['r_delim']
            self.skip_header = settings['skip_header']

            s.INPUT_FILE = 'test/input.csv'
            s.OUTPUT_FILE = 'test/output.docx'
            s.NUM_NON_BODY_COLS = 5
            
            s.ID_IND = 0
            s.HEADING_LEVEL_IND = 2
            s.HEADING_NUM_IND = 3
            s.HEADING_TEXT_IND = 4
            s.BODY_TEXT_IND = 5

            s.title    = 'Payload Specification'
            s.subject  = 'Auto export using docx from Python'
            s.creator  = 'Randy Sage'
            s.keywords = ['python', 'Office Open XML', 'Word']

            return s
        except 23:
            sys.exit("Unexpected issue in %s\nExiting..." %
                     inspect.stack()[0][3])
    # read_json

# MySettings


class DocxConfig():
    def __init__(self, settings):
        self.s = settings

        # Default set of relationshipships - the minimum components of a document
        self.relationships = relationshiplist()

        # Make a new document tree - this is the main part of a Word document
        self.document = newdocument()

        # This xpath location is where most interesting content lives
        self.body = self.document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    # end __init__

    def add_image(self, image_file, image_caption):
        self.relationships, picpara = picture(self.relationships, 
                                              image_file,
                                              image_caption)
        self.body.append(picpara)        
    # end add_image

    def save(self, out_file):
        s = self.s
        # Create our properties, contenttypes, and other support files
    
        coreprops = coreproperties(title = s.title, 
                                   subject = s.subject, 
                                   creator = s.creator,
                                   keywords = s.keywords)
        appprops = appproperties()
        content_types = contenttypes()
        web_settings = websettings()
        word_relationships = wordrelationships(self.relationships)

        #print 'About to save'
        # Save our document
        savedocx(self.document, coreprops, appprops, content_types, web_settings,
                 word_relationships, out_file)
    # end save

    def write_heading(self, heading_text, heading_level):
        self.body.append(heading(heading_text, 
                                 heading_level))

    def write_paragraph(self, para, row_id):
        try:
            self.body.append(paragraph(para))
        except:
            self.body.append(paragraph(
                'Failed to write paragraph with id %s' % row_id))
    # end write_paragraph
# DocxConfig

class CsvParser():
    def __init__(self, settings):
        self.s = settings
        # s.l_delim
        # s.r_delim
        self.out_docx = None # Error if not set; TODO improve error handling
    # end __init__

    def handle_token(self, token):
        # Add an image - INITIAL HACK (but found/fixed a bug)
        self.out_docx.add_image( 'test/images/240px-Smiley.svg.png',
                                 'Captions not implemented - TBD')
    # end handle_token

    def output_body_to_docx(self, body, row_id):
        non_matches = re.split('{[^}]*}',body)
        matches = re.findall('{[^}]*}',body)
        if len(non_matches) != len(matches)+1:
            raise Exception( "Need an error to throw...  todo")

        for i in range(0,len(matches)):
            self.out_docx.write_paragraph(non_matches[i],row_id)
            self.handle_token(matches[i])
        self.out_docx.write_paragraph(non_matches[-1],row_id)
    # end output_body_to_docx

    def output_row_to_docx(self, row):
        s = self.s

        if not len(row): # no content
            return

        if len(row[s.HEADING_LEVEL_IND]):
            h_text = ' '.join((row[s.HEADING_NUM_IND],
                               row[s.HEADING_TEXT_IND]))
            self.out_docx.write_heading(h_text, 
                                   int(row[s.HEADING_LEVEL_IND]))
        else:
            self.output_body_to_docx(row[s.BODY_TEXT_IND],row[s.ID_IND])
    # end output_row_to_docx

    def parse(self):
        if self.out_docx == None:
            sys.exit('Output docx not configured\nExiting...' )
        try:
            with open(self.s.INPUT_FILE,'rb') as csvfile:
                reader = csv.reader(csvfile, delimiter=',', quotechar='"')
                need_to_skip_header = self.s.skip_header
                for row in reader:
                    if need_to_skip_header:
                        # Do nothing when skipping, except no longer skip
                        need_to_skip_header = False
                    else:
                        self.output_row_to_docx(row)
        except IOError:
            sys.exit('Failed to open input file: %s\nExiting...' % 
                     self.s.INPUT_FILE)
        # end try
    # end parse()

    def write_docx(self, out_docx):
        self.out_docx = out_docx
        self.parse()
    # end write_docx()
#end CsvParser
    
if __name__ == '__main__':


    s = MySettings()
    s.read_json()

    out_docx = DocxConfig(s)
    csv_parser = CsvParser(s)
    
    csv_parser.write_docx(out_docx)

    out_docx.save(s.OUTPUT_FILE)
    print 'Done :-)'

#end __main__
