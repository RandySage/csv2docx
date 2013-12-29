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
 - Todo items will be placed throughout text preceded by 'TODO' with a colon, to cue eclipse
     recommend `grep -rI TODO [^v]*` if not using eclipse (exclude virtualenv)

"""


from docx import *
import argparse
import re
import csv
import sys
import json
import inspect
import logging

FORMAT = '%(asctime)-15s %(module)s %(funcName)s %(message)s'
logging.basicConfig(format=FORMAT, filename='temp.log')
log = logging.getLogger()

DEFAULT_JSON = 'test/test_settings.json'
DEFAULT_INPUT_FILE = 'test/input.csv'
DEFAULT_OUTPUT_FILE = 'test/output.docx'

class LogicError(Exception):
    pass
class JsonError(Exception):
    pass
class CrossRefError(Exception):
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

class utils():
    
    @staticmethod
    def log(msg, logfunction=log.warning, ex=None):
        print msg
        logfunction(msg)
        if ex:
            log.info(ex)
    @staticmethod        
    def int_repr(s):
        try:
            return int(s)
        except ValueError:
            return None


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

        except Exception as ex:
            sys.exit("Unexpected issue in %s\nExiting..." %
                     inspect.stack()[0][3])
            raise # Previous line should raise - here for pcregrep audit
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
    
    def valid_character(self, i):
        if not isinstance(i, int) and isinstance(i,str):
            if len(i) > 1:
                raise LogicError('%s was passed a string with length > 1' %
                                 inspect.stack()[0][3])
            i = ord(i)  
        return ( # conditions ordered by presumed frequency
            0x20 <= i <= 0xD7FF 
            or i in (0x9, 0xA, 0xD)
            or 0xE000 <= i <= 0xFFFD
            or 0x10000 <= i <= 0x10FFFF
            )
    # end valid_character

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

    def clean(self, text):
        try:
            clean_text = ''.join(
                unicode(c,encoding='ascii',errors='ignore') for c in text
                ) 
            return clean_text
        except Exception as ex:
            print ('WARNING: Unexpected error encountered in %s'  %
                   inspect.stack()[0][3])
            raise # Don't catch all without raising
        # end clean

    def write_heading(self, heading_text, heading_level):
        clean_text = self.clean(heading_text)
        self.body.append(heading(clean_text,
                                 heading_level))

    def write_paragraph(self, para, row_id):
        try:
            if len(para):
                # TODO: implement separate paragraphs where newlines were included
                # see http://stackoverflow.com/a/14422406/527489
                clean_para = self.clean(para)
                self.body.append(paragraph(clean_para))
        except Exception as ex:
            err_msg = 'Failed to write paragraph with id %s' % row_id
            self.body.append(paragraph(err_msg))
            log.warning(err_msg)
            log.info(ex)
            raise # Don't catch all without raising
    # end write_paragraph
# DocxConfig

class CsvParser():
    
    def __init__(self, settings):
        self.header_dict = {}
        self.s = settings
        self.out_docx = None # Error if not set; TODO improve error handling
    # end __init__

    class ParsedToken:
        is_image = False
        value = ''
    # end ParsedToken

    def insert_image(self, filename_or_other, id=-1):
        # Add an image - INITIAL HACK (but found/fixed a bug)
        try:
            self.out_docx.add_image( filename_or_other,
                                     'Captions not implemented - TBD')
        except IOError as ex:
            log.exception(str(ex))
        except Exception as ex:
            print "Assumed to be a block of text in delimeters" ,
            self.out_docx.write_paragraph('{'+filename_or_other+'}', id)
            print "."
            # TODO: do something with tables
            # TODO: finish getting images to work (done?)
            raise # Don't catch all without raising
 
    # end insert_image

    def parse_token(self, token):
        s = self.s
        token_contents = token[len(s.l_delim):-len(s.r_delim)]
        parsed = self.ParsedToken()

        if re.match('^[#H]\d+$',token_contents):
            try:
                if (token_contents[0:len(s.heading_text_symbol)] == 
                    s.heading_text_symbol):
                    dict_index = s.HEADING_TEXT_IND
                elif (token_contents[0:len(s.heading_number_symbol)] == 
                    s.heading_number_symbol):
                    dict_index = s.HEADING_NUM_IND
                else:
                    print "OOPS: %s" % token_contents[0] 
                target_key = int(token_contents[1:])
                if not self.header_dict.has_key(target_key):
                    raise CrossRefError('Did not find cross reference key, %d' % target_key)
                target_dict = self.header_dict[target_key]
                parsed.value = target_dict[dict_index]
                parsed.is_image = False
                return parsed
            except Exception as ex:
                log.error("%s caught exception processing token, %s" %
                          (inspect.stack()[0][3], token))
                raise # Don't catch all without raising
        else:
            log.info('IMAGE OR NOTE: %s' % token_contents)
            parsed.value = token_contents
            parsed.is_image = True
            return parsed
    # end parse_token

    def replace_cross_refs(self, body, row_id):
        result = []
        try:
            s = self.s
            esc_seq = '%s[^%s]*%s' % ( s.l_delim, s.r_delim, s.r_delim )
            non_matches = re.split(esc_seq,body)
            matches = re.findall(esc_seq,body)
            if len(non_matches) != len(matches)+1:
                raise LogicError('%s erred in regex logic' %
                                 inspect.stack()[0][3])
    
            this_str = ''
            for i in range(0,len(matches)):
                this_str += non_matches[i]
                parsed = self.parse_token(matches[i]) 
                if parsed.is_image:
                    result.append(this_str)
                    this_str = ''
                    result.append(parsed)
                else:
                    this_str += parsed.value
            result.append(this_str)
            result.append(non_matches[-1])
        except Exception as ex:
            print "Warning: did not write this data...\n%s" % str(row_id)+": "+body
            raise # Don't catch all without raising
        #end try
        return result
    # end replace_cross_refs

    def output_body_to_docx(self, body, row_id):
        replaced_list = self.replace_cross_refs(body, row_id)
        for entry in replaced_list:
            if isinstance(entry, Parsed):
                self.insert_image(entry.value)
            else:
                self.out_docx.write_paragraph(entry, row_id)
            # endif
        # endfor
    # end output_body_to_docx

    def output_header_to_docx(self, row):
        try:
            h_text = ' '.join((row[s.HEADING_NUM_IND],
                               row[s.HEADING_TEXT_IND]))
            self.out_docx.write_heading(h_text, 
                                        int(row[s.HEADING_LEVEL_IND]))
        except (SystemError, SystemExit):
            raise
        except Exception as ex:
            print "Warning: did not write this heading...\n%s" % row
            raise # Don't catch all without raising

    def output_row_to_docx(self, row):
        s = self.s

        if not len(row): # no content
            return

        if len(row[s.HEADING_LEVEL_IND]):
            self.output_header_to_docx(row)
        else:
            self.output_body_to_docx(row[s.BODY_TEXT_IND],row[s.ID_IND])
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

    def build_header_dict(self):
        s = self.s
        header_dict = {}
        with open(s.INPUT_FILE,'U') as csvfile:
            reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            for row in reader:
                if not len(''.join(row)) or (s.ID_IND >= len(row)):
                    utils.log("row has fewer than %d entries\nRow: %s" % 
                               (s.ID_IND+1,row))
                else:
                    int_key = utils.int_repr(row[s.ID_IND])
                    if int_key == None or header_dict.has_key(int_key):
                        utils.log('WARNING: Non-int or dupl key - ignoring: %s' % 
                                  row[s.ID_IND])
                    else:
                        clean_row = self.clean_backslash_r(row)
                        header_dict[int_key] = {}
                        header_dict[int_key][s.HEADING_NUM_IND] = (
                            clean_row[s.HEADING_NUM_IND] )
                        header_dict[int_key][s.HEADING_TEXT_IND] = (
                            clean_row[s.HEADING_TEXT_IND])
                #end if/else
            #end for    
        #end with    
        return header_dict
    #end build_header_dict
    
    def parse(self):
        s = self.s
        if self.out_docx == None:
            sys.exit('Output docx not configured\nExiting...' )
        try:
            self.header_dict = self.build_header_dict()
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
