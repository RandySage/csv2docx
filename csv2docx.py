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
from curses import ascii
import fnmatch

THIS_FOLDER = os.path.abspath(os.path.dirname(__file__))
FORMAT = '%(asctime)-15s %(module)s %(funcName)s %(message)s'
logging.basicConfig(format=FORMAT,
                    filename=os.path.join(THIS_FOLDER, 'temp.log'))
log = logging.getLogger()

DEFAULT_JSON = 'test/test_settings.json'
DEFAULT_INPUT_FILE = 'test/input.csv'
DEFAULT_OUTPUT_FILE = 'test/output.docx'

def valid_XML_char(c):
    i = ord(c)
    return (# conditions ordered by presumed frequency
        0x20 <= i <= 0xD7FF
        or i in (0x9, 0xA, 0xD)
        or 0xE000 <= i <= 0xFFFD
        or 0x10000 <= i <= 0x10FFFF
        )

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
                        default=DEFAULT_INPUT_FILE)
    parser.add_argument("--output", "-o",
                        help='output docx filename',
                        default=DEFAULT_OUTPUT_FILE)
    parser.add_argument("--settings", "-s",
                        help='settings file (json)',
                        default=DEFAULT_JSON)

    return parser

class utils():

    @staticmethod
    def log(msg, ex=None, logfunction=log.warning):
        print msg
        logfunction(msg)
        if ex:
            log.info(ex)
    # end log


    @staticmethod
    def int_repr(s):
        try:
            return int(s)
        except ValueError:
            return None
    # end int_repr

    @staticmethod
    def locator(pattern, root=os.curdir):
        """Locate all files matching supplied filename pattern in and below
        supplied root directory."""
        for path, dirs, files in os.walk(os.path.abspath(root)):
            for filename in fnmatch.filter(files, pattern):
                yield os.path.join(path, filename)
    # end locate


class MySettings():
    """ Container class for csv2docx settings """

    id_ind = None
    heading_level_ind = None
    heading_num_ind = None
    heading_text_ind = None
    body_text_ind = None
    all_inds = []

    def json_file_to_dict(self, json_filename):
        try:
            with open(json_filename) as json_file:
                return json.loads(json_file.read())
        except ValueError:
            sys.exit("Parse error for json file: %s\nExiting..." % json_filename)
        except IOError:
            sys.exit("Exiting from %s because failed to open json file: %s\nExiting..." %
                     (inspect.stack()[0][3], json_filename))
    # end def

    def validate_set_json_dict(self, json_dict):
        try:
            for k, v in json_dict.items():
                setattr(self, k, v)

            if (len(self.l_delim) > 1 or
                 len(self.r_delim) > 1):
                sys.exit(
                    ("l_delim ('%s') and r_delim ('%s') need to be single characters\nExiting..." %
                     (settings['l_delim'], settings['r_delim'])))

            self.all_inds.append(self.id_ind)
            self.all_inds.append(self.heading_level_ind)
            self.all_inds.append(self.heading_num_ind)
            self.all_inds.append(self.heading_text_ind)
            self.all_inds.append(self.body_text_ind)
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
        if not isinstance(i, int) and isinstance(i, str):
            if len(i) > 1:
                raise LogicError('%s was passed a string with length > 1' %
                                 inspect.stack()[0][3])
            i = ord(i)
        return (# conditions ordered by presumed frequency
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

        coreprops = coreproperties(title=s.title,
                                   subject=s.subject,
                                   creator=s.creator,
                                   keywords=s.keywords)
        appprops = appproperties()
        content_types = contenttypes()
        web_settings = websettings()
        word_relationships = wordrelationships(self.relationships)

        # print 'About to save'
        # Save our document
        savedocx(self.document, coreprops, appprops, content_types, web_settings,
                 word_relationships, out_file)
    # end save

#     import re
#     # match characters from [upside down question mark] to the end of the JSON-encodable range
#     @staticmethod
#     def isprintable(s):
#         exclude = re.compile(ur'[\u00bf-\uffff]')
#         return not bool(exclude.search(s))

    @staticmethod
    def clean(text):
        try:
            clean_text = ''.join(
                (valid_XML_char(c) and c) or
                ((c == '\n' or c == '\r') and '\n') or
                '?' for c in text
                )
                # curses.unctrl(c) for c in text
                # (DocxConfig.isprintable(c) and c) or
            return clean_text
        except Exception as ex:
            print ('WARNING: Unexpected error encountered in %s' %
                   inspect.stack()[0][3])
            raise # Don't catch all without raising
    # end clean

    def write_heading(self, heading_text, heading_level):
        self.body.append(heading(heading_text,
                                 heading_level))

    def write_paragraph(self, para, row_id):
        try:
            if len(para):
                # TODO: implement separate paragraphs where newlines were included
                # see http://stackoverflow.com/a/14422406/527489
                for para_text in para.split('\n'):
                    self.body.append(paragraph(para_text))
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
        self.s = settings
        self.out_docx = None # Error if not set; TODO improve error handling
        self.build_clean_dict()
    # end __init__

    def get_clean_dict(self):
        """ Return dictionary of header references """
        return self.clean_dict
    # get_clean_dict

    class ParsedToken:
        is_image = False
        value = ''
        def __repr__(self):
            return "<is_image=%s, value=%s>" % (self.is_image, self.value)
    # end ParsedToken

    def insert_image(self, filename_or_other, id=-1):
        # Add an image - INITIAL HACK (but found/fixed a bug)
        try:
            self.out_docx.add_image(filename_or_other,
                                     'Captions not implemented - TBD')
        except IOError as ex:
            log.exception(str(ex))
        except Exception as ex:
            print "Assumed to be a block of text in delimeters" ,
            self.out_docx.write_paragraph('{' + filename_or_other + '}', id)
            print "."
            # TODO: do something with tables
            # TODO: finish getting images to work (done?)
            raise # Don't catch all without raising

    # end insert_image

    def parse_token(self, token):
        s = self.s
        token_contents = token[len(s.l_delim):-len(s.r_delim)]
        parsed = self.ParsedToken()

        if re.match('^[#H]\d+$', token_contents):
            try:
                if (token_contents[0:len(s.heading_text_symbol)] ==
                    s.heading_text_symbol):
                    dict_index = s.heading_text_ind
                elif (token_contents[0:len(s.heading_number_symbol)] ==
                    s.heading_number_symbol):
                    dict_index = s.heading_num_ind
                else:
                    print "OOPS: %s" % token_contents[0]
                target_key = int(token_contents[1:])
                if not self.clean_dict.has_key(target_key):
                    print repr(self.clean_dict)
                    raise CrossRefError('Did not find cross reference key, %d' % target_key)
                target_dict = self.clean_dict[target_key]
                parsed.value = target_dict[dict_index]
                parsed.is_image = False
                return parsed
            except Exception as ex:
                utils.log("%s caught exception processing token, %s" %
                          (inspect.stack()[0][3], token), ex=ex)
                raise # Don't catch all without raising
        else:
            log.info('IMAGE OR NOTE: %s' % token_contents)
            parsed.value = token_contents
            parsed.is_image = True
            return parsed
    # end parse_token

    def replace_tokens(self, body, row_id):
        result = []
        try:
            s = self.s
            esc_seq = '%s[^%s]*%s' % (s.l_delim, s.r_delim, s.r_delim)
            non_matches = re.split(esc_seq, body)
            matches = re.findall(esc_seq, body)
            if len(non_matches) != len(matches) + 1:
                raise LogicError('%s erred in regex logic' %
                                 inspect.stack()[0][3])

            this_str = ''
            for i in range(0, len(matches)):
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
            print "Warning: did not write this data...\n%s" % str(row_id) + ": " + body
            raise # Don't catch all without raising
        # end try
        return result
    # end replace_tokens

    def output_body_to_docx(self, body, row_id):
        replaced_list = self.replace_tokens(body, row_id)
        for entry in replaced_list:
            if isinstance(entry, CsvParser.ParsedToken):
                self.insert_image(entry.value)
            else:
                self.out_docx.write_paragraph(entry, row_id)
            # endif
        # endfor
    # end output_body_to_docx

    def output_header_to_docx(self, row):
        s = self.s
        try:
            h_text = ' '.join((row[s.heading_num_ind],
                               row[s.heading_text_ind]))
            self.out_docx.write_heading(h_text,
                                        int(row[s.heading_level_ind]))
        except (SystemError, SystemExit):
            raise
        except Exception as ex:
            print "Warning: did not write this heading...\n%s" % row
            raise # Don't catch all without raising

    def write_debug_csv_data(self, row, debug_writer):
        s = self.s
        row_list = [row[ind] for ind in (s.id_ind,
                                         s.heading_level_ind,
                                         s.heading_num_ind,
                                         s.heading_text_ind,
                                         s.body_text_ind)]
        replaced_body = repr(self.replace_tokens(
                                row[s.body_text_ind], row[s.id_ind]))
        row_list.append(replaced_body)
        debug_writer.writerow(row_list)
    # end write_debug_csv_data

    def output_row_to_docx(self, row_id, debug_writer=None):
        s = self.s
        row = self.clean_dict[row_id]
        if len(row[s.heading_level_ind]):
            self.output_header_to_docx(row)
        else:
            self.output_body_to_docx(row[s.body_text_ind], row[s.id_ind])
        if debug_writer:
            self.write_debug_csv_data(row, debug_writer)
    # end output_row_to_docx

    def clean_only(self, row):
        s = self.s
        new_row = row[:]
        for i in s.all_inds:
            new_row[i] = DocxConfig.clean(row[i])
            # new_row[i] = '\n'.join((DocxConfig.clean(line) for line in row[i].split('\n')))
            # TODO: clean up previous line
        return new_row
    # clean_only

    def clean_n_parse_tokens(self, row, debug=False):
        s = self.s
        new_row = self.clean_only(row)
        for i in s.all_inds:
            new_row[i] = self.replace_tokens(new_row[i], new_row[s.id_ind])
            # TODO: this is not finished
            # new_row[i] = DocxConfig.clean(row[i])
        return new_row
    # end clean_n_parse_tokens

    def build_clean_dict(self):
        """Builds a dictionary representation of input csv"""
        s = self.s
        self.clean_dict = {}
        self.ordered_id_list = []
        skipped_header = False
        with open(s.INPUT_FILE, 'rb') as csvfile:
            reader = csv.reader(csvfile, delimiter=',', quotechar='"')
            for row in reader:
                if not len(''.join(row)) or (s.id_ind >= len(row)):
                    utils.log("row has fewer than %d entries\nRow: %s" %
                               (s.id_ind + 1, row))
                elif skipped_header:
                    int_key = utils.int_repr(row[s.id_ind])
                    if int_key == None or self.clean_dict.has_key(int_key):
                        utils.log('WARNING: Non-int or dupl key - ignoring extras: %s' %
                                  row[s.id_ind])
                    else:
                        self.ordered_id_list.append(int_key)
                        self.clean_dict[int_key] = self.clean_only(row)
                else:
                    skipped_header = True
    # end build_clean_dict

    def write_docx(self, out_docx, debug=False):
        self.out_docx = out_docx
        with open('debug.csv', 'wb') as debug_file:
            debug_writer = csv.writer(debug_file)
            for row_id in self.ordered_id_list:
                self.output_row_to_docx(row_id, debug_writer)
        self.out_docx.write_paragraph(unichr(10146), None)
    # end write_docx()

# end CsvParser

if __name__ == '__main__':
    parser = create_parser()
    args = parser.parse_args()

    s = MySettings()
    s.read_json_file(args.settings) # Use argparse specified settings file

    # Add the argparse inputs
    s.INPUT_FILE = args.input
    s.OUTPUT_FILE = args.output

    debug = hasattr(s, 'debug') and s.debug


    out_docx = DocxConfig(s)
    csv_parser = CsvParser(s)

    csv_parser.write_docx(out_docx, debug=debug)

    out_docx.save(s.OUTPUT_FILE)
    print 'Done :-)'

# end __main__
