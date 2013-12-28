#! /usr/bin/env python

"""
This file s based off the techniques from the example for python-docx --
 http://github.com/mikemaccana/python-docx
and uses that module to implement a docx generator.  Currently it operates on 
a csv, but there is an attempt to make it modular enough that text data could 
come from other sources.

This file makes a .docx (Word 2007) file from scratch.

See LICENSE for licensing information.

Todo:
 - do something with tables
 - finish getting images to work
 - deal with line endings in cells
 - fix to use l_delim, r_delim

"""


from docx import *
import argparse
import re
import csv
import sys
import json
import inspect

DEFAULT_JSON = 'settings.json'
DEFAULT_INPUT_FILE = 'test/input.csv'
DEFAULT_OUTPUT_FILE = 'test/output.docx'

class LogicError(Exception):
    pass
class JsonError(Exception):
    pass

def create_parser():
    parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    
    parser.add_argument("--input", "-i",
                        help="input csv filename",
                        default = DEFAULT_INPUT_FILE)
    parser.add_argument("--output", "-o",
                        help='output docx filename',
                        default = DEFAULT_OUTPUT_FILE)
    parser.add_argument("--settings", "-s",
                        help='settings file (json)',
                        default = DEFAULT_JSON)

    return parser


class MySettings():

    def json_file_to_dict(self, json_filename):
        try:
            with open(json_filename) as json_file:
                return json.loads(json_file.read())
        except ValueError:
            sys.exit("Parse error for json file: %s\nExiting..." % json_filename)        
        except IOError:
            sys.exit("Exiting from %s because failed to open json file: %s\nExiting..." % 
                     (inspect.stack()[0][3],json_filename))
    # end def
    
    def validate_set_json_dict(self, json_dict):
        try:
            for k, v in json_dict.items():
                setattr(self, k, v)

            if ( len(self.l_delim) > 1 or
                 len(self.r_delim) > 1 ):
                sys.exit(
                    ("l_delim ('%s') and r_delim ('%s') need to be single characters\nExiting..." % 
                     (settings['l_delim'], settings['r_delim'])))

            # Confirm settings about as expected

        except:
            sys.exit("Unexpected issue in %s\nExiting..." %
                     inspect.stack()[0][3])
    # validate_set_json_dict

    def read_json_file(self, json_file):
        json_dict = self.json_file_to_dict(json_file)
        self.validate_set_json_dict(json_dict)
    # read_json_file
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
        self.header_dict = {}
        self.s = settings
        self.out_docx = None # Error if not set; TODO improve error handling
    # end __init__

    def insert_image(self, filename_or_other):
        # Add an image - INITIAL HACK (but found/fixed a bug)
        try:
            self.out_docx.add_image( filename_or_other,
                                     'Captions not implemented - TBD')
        except:
            self.out_docx.write_paragraph(filename_or_other)
    # end insert_image

    def parse_token(self, token):
        s = self.s
        token_contents = token[len(s.l_delim):-len(s.r_delim)]

        if re.match('^[#H]\d+$',token_contents):
            try:
                if token_contents[0] == '#':
                    dict_index = 0
                elif token_contents[0] == 'H':
                    dict_index = 1
                else:
                    print "OOPS: %s" % token_contents[0] 
                print token_contents[1:] + " " ,
                target_key = int(token_contents[1:])
                target_dict = self.header_dict[target_key]
                return target_dict[dict_index]
            except:
                print "except"
                return (False, token + "_ERROR")
        else:
            print 'IMAGE OR NOTE: %s' % token_contents
            return (True, token_contents)
    # end parse_token

    def output_body_to_docx(self, body, row_id):
        s = self.s
        esc_seq = '%s[^%s]*%s' % ( s.l_delim, s.r_delim, s.r_delim )
        non_matches = re.split(esc_seq,body)
        matches = re.findall(esc_seq,body)
        if len(non_matches) != len(matches)+1:
            raise LogicError('%s erred in regex logic' %
                             (inspect.stack()[0][3],json_filename))

        this_str = ''
        for i in range(0,len(matches)):
            this_str += non_matches[i]
            is_image, ref_str = self.parse_token(matches[i])
            if is_image:
                self.out_docx.write_paragraph(this_str, row_id)
                this_str = ''
                self.insert_image(ref_str)
            else:
                this_str += ref_str
        self.out_docx.write_paragraph(non_matches[-1],row_id)
    # end output_body_to_docx

    def output_row_to_docx(self, row):
        s = self.s

        if not len(row): # no content
            return

        try:
            if len(row[s.HEADING_LEVEL_IND]):
                h_text = ' '.join((row[s.HEADING_NUM_IND],
                                   row[s.HEADING_TEXT_IND]))
                self.out_docx.write_heading(h_text, 
                                            int(row[s.HEADING_LEVEL_IND]))
            else:
                self.output_body_to_docx(row[s.BODY_TEXT_IND],row[s.ID_IND])
        except:
            print "Warning: did not write this data...\n%s" % row
    # end output_row_to_docx
    
    def clean_backslash_r(self, row, debug=False):
        s = self.s
        new_row = [unicode(col,errors='ignore') for col in row]
        #try:
        if( hasattr(s,'indices_to_replace_backslash_r') and 
            len(s.indices_to_replace_backslash_r) and
            len(row) # rule out empty rows
        ):
            if not hasattr(s,'replace_backslash_r_with'):
                raise JsonError('json has indices_to_replace_backslash_r ' + 
                                'true-like, but no replace_backslash_r_with')
            for index in s.indices_to_replace_backslash_r:
                if not isinstance(index, int) or abs(index) > len(new_row):
                    raise JsonError(('json has indices_to_replace_backslash_r ' + 
                                     'with value %s not in row %s') %
                                    (index, row[s.ID_IND])
                    )
                else:
                    new_row[index] = new_row[index].replace('\r', 
                                                            s.replace_backslash_r_with)
            # end replace for
        return new_row
    # end clean_backslash_r

    def parse(self):
        s = self.s
        if self.out_docx == None:
            sys.exit('Output docx not configured\nExiting...' )
        try:
            with open(s.INPUT_FILE,'U') as csvfile:
                reader = csv.reader(csvfile, delimiter=',', quotechar='"')
                for row in reader:
                    if len(row):
                        try:
                            row = self.clean_backslash_r(row)
                            int_key = int(row[s.ID_IND])
                            self.header_dict[int_key] = (row[s.HEADING_NUM_IND],
                                                         row[s.HEADING_TEXT_IND])
                        except:
                            print "issue in entry with id %s" % row[self.s.ID_IND]
                            pass
            with open(s.INPUT_FILE,'rb') as csvfile:
                reader = csv.reader(csvfile, delimiter=',', quotechar='"')
                need_to_skip_header = s.skip_header
                for row in reader:
                    if need_to_skip_header:
                        # Do nothing when skipping, except no longer skip
                        need_to_skip_header = False
                    else:
                        self.output_row_to_docx(row)
        except IOError:
            sys.exit('Failed to open input file: %s\nExiting...' % 
                     s.INPUT_FILE)
        # end try
    # end parse()

    def write_docx(self, out_docx):
        self.out_docx = out_docx
        self.parse()
    # end write_docx()
#end CsvParser
    
if __name__ == '__main__':
    parser = create_parser()
    args = parser.parse_args()

    s = MySettings()
    s.read_json_file(args.settings) # Use argparse specified settings file
    
    # Add the argparse inputs
    s.INPUT_FILE = args.input
    s.OUTPUT_FILE = args.output

    out_docx = DocxConfig(s)
    csv_parser = CsvParser(s)
    
    csv_parser.write_docx(out_docx)

    out_docx.save(s.OUTPUT_FILE)
    print 'Done :-)'

#end __main__
