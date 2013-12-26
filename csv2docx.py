#! /usr/bin/env python

"""
This file makes a .docx (Word 2007) file from scratch, showing off most of the
features of python-docx.

If you need to make documents from scratch, you can use this file as a basis
for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

from docx import *
import re
import csv
import sys
import json
import inspect

DEFAULT_JSON = 'settings.json'
INPUT_FILE = 'test/input.csv'
OUTPUT_FILE = 'test/output.docx'
NUM_NON_BODY_COLS = 5

ID_IND = 0
HEADING_LEVEL_IND = 2
HEADING_NUM_IND = 3
HEADING_TEXT_IND = 4
BODY_TEXT_IND = 5

title    = 'Payload Specification'
subject  = 'Auto export using docx from Python'
creator  = 'Randy Sage'
keywords = ['python', 'Office Open XML', 'Word']


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
            return s
        except 23:
            sys.exit("Unexpected issue in %s\nExiting..." %
                     inspect.stack()[0][3])
    # read_json

# MySettings


class DocxConfig():
    # Default set of relationshipships - the minimum components of a document
    relationships = relationshiplist()

    # Make a new document tree - this is the main part of a Word document
    document = newdocument()

    # This xpath location is where most interesting content lives
    body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

    def add_image(self, image_file, image_caption):
        self.relationships, picpara = picture(relationships, 
                                                   image_file,
                                                   image_caption)
        self.body.append(picpara)        
    # end add_image

    def save(self, out_file):
        # Create our properties, contenttypes, and other support files
    
        coreprops = coreproperties(title=title, subject=subject, creator=creator,
                               keywords=keywords)
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
    
if __name__ == '__main__':

    s = MySettings()
    s.read_json()
    print s.l_delim
    print s.r_delim

    out_docx = DocxConfig()

    # input_file = open('export.csv')
    # input_dat = input_file.read()
    # input_file.close()

    def handle_token(token):
        # Add an image
        out_docx.add_image( 'image1.png',
                            'Captions not implemented - TBD')
    # end handle_token

    def output_body_to_docx(body, row_id):
        non_matches = re.split('{[^}]*}',body)
        matches = re.findall('{[^}]*}',body)
        if len(non_matches) != len(matches)+1:
            raise Exception( "Need an error to throw...  todo")

        for i in range(0,len(matches)):
            out_docx.write_paragraph(non_matches[i],row_id)
            handle_token(matches[i])
        out_docx.write_paragraph(non_matches[-1],row_id)
    # end output_body_to_docx

    def output_row_to_docx(row):

        if len(row[HEADING_LEVEL_IND]):
            h_text = ' '.join((row[HEADING_NUM_IND],
                               row[HEADING_TEXT_IND]))
            out_docx.write_heading(h_text, 
                                   int(row[HEADING_LEVEL_IND]))
        else:
            output_body_to_docx(row[BODY_TEXT_IND],row[ID_IND])
    # end output_row_to_docx

    try:
        with open(INPUT_FILE,'rb') as csvfile:
            reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            already_skipped_header = False
            for row in reader:
                if already_skipped_header:
                    output_row_to_docx(row)
                else:
                    already_skipped_header = True
    except IOError:
        sys.exit('Failed to open input file: %s\nExiting...' % INPUT_FILE)

    # for this_line in input_lines[1:7]:  # Skip heading line
    #     #cols = re.findall("[^,]*",this_line)
    #     non_body_regex = r"^(?:[^,]*,){%d}" % NUM_NON_BODY_COLS
    #     non_body_col_text = ''.join(re.findall(non_body_regex, this_line))
    #     print "Pre-body (%d): %s" % (len(non_body_col_text), non_body_col_text)

    #     body_text = this_line[len(non_body_col_text):]
        
    #     strip_commas_quotes_re = ',*"(.*)",*'
    #     stripped_body_text = re.findall(strip_commas_quotes_re, body_text)
    #     if len(stripped_body_text):
    #         body_text = stripped_body_text[0]
    #     print body_text

    # print '====== DONE ======\n'*3
    # Append two headings and a paragraph
    # body.append(heading("Welcome to Python's docx module", 1))
    # body.append(heading('Make and edit docx in 200 lines of pure Python', 2))
    # body.append(paragraph('The module was created when I was looking for a '
    #     'Python support for MS Word .doc files on PyPI and Stackoverflow. '
    #     'Unfortunately, the only solutions I could find used:'))

    # Add a numbered list
    points = [ 'COM automation'
             , '.net or Java'
             , 'Automating OpenOffice or MS Office'
             ]
    # for point in points:
    #     body.append(paragraph(point, style='ListNumber'))
    # body.append(paragraph([('For those of us who prefer something simpler, I '
    #                       'made docx.', 'i')]))    
    # body.append(heading('Making documents', 2))
    # body.append(paragraph('The docx module has the following features:'))

    # # Add some bullets
    # points = ['Paragraphs', 'Bullets', 'Numbered lists',
    #           'Multiple levels of headings', 'Tables', 'Document Properties']
    # for point in points:
    #     body.append(paragraph(point, style='ListBullet'))

    # body.append(paragraph('Tables are just lists of lists, like this:'))
    # # Append a table
    # tbl_rows = [ ['A1', 'A2', 'A3']
    #            , ['B1', 'B2', 'B3']
    #            , ['C1', 'C2', 'C3']
    #            ]
    # body.append(table(tbl_rows))

    # body.append(heading('Editing documents', 2))
    # body.append(paragraph('Thanks to the awesomeness of the lxml module, '
    #                       'we can:'))
    # points = [ 'Search and replace'
    #          , 'Extract plain text of document'
    #          , 'Add and delete items anywhere within the document'
    #          ]
    # for point in points:
    #     body.append(paragraph(point, style='ListBullet'))

    # # Search and replace
    # print 'Searching for something in a paragraph ...',
    # if search(body, 'the awesomeness'):
    #     print 'found it!'
    # else:
    #     print 'nope.'

    # print 'Searching for something in a heading ...',
    # if search(body, '200 lines'):
    #     print 'found it!'
    # else:
    #     print 'nope.'

    # print 'Replacing ...',
    # body = replace(body, 'the awesomeness', 'the goshdarned awesomeness')
    # print 'done.'

    # # Add a pagebreak
    # body.append(pagebreak(type='page', orient='portrait'))

    # body.append(heading('Ideas? Questions? Want to contribute?', 2))
    # body.append(paragraph('Email <python.docx@librelist.com>'))


    out_docx.save(OUTPUT_FILE)
    print 'Done :-)'

