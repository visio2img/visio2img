#!/usr/bin/env python3

import win32com.client
from pywintypes import com_error
from win32com.client import constants
from os import path, chdir, getcwd

from sys import exit, stderr
from optparse import OptionParser

__all__ = ['export_img']

class IllegalImageFormatException(TypeError):
    """
    This exception means Exceptions for Illegal Image Format.
    """

def get_dispatch_format(extension):
    return 'Visio.InvisibleApp' # vsd format


def get_pages(app, page_num=None):
    """
    app -> page
    if page_num is None, return all pages.
    if page_num is int object, return path_num-th page(from 1).
    """
    pages = app.ActiveDocument.Pages
    return [list(pages)[page_num - 1]] if page_num else pages

def export_img(in_filename, out_filename, page_num=None):
    """
    export as image format
    """
    # to absolute path
    in_filename = path.abspath(in_filename)
    out_filename = path.abspath(out_filename)
    
    # define filename without extension and extension variable
    in_filename_without_extension, in_extension = path.splitext(in_filename)
    out_filename_without_extension, out_extension = path.splitext(out_filename)

    if in_extension not in ('.vsd'):
        err_str = (
                'Input filename is not llegal for visio file. \n' 
                'This program is suppert only vsd extension.'
                )
        raise IllegalImageFormatException(err_str)

    if out_extension not in ('.gif', '.jpg', '.jpeg', '.png'):
                err_str = (
                'Output filename is not llegal for visio file. \n' 
                'This program is suppert gif, jpeg, png extension.'
                )
                raise IllegalImageFormatException(err_str)

    # if file is not found, exit from program
    if not path.exists(in_filename):
        raise FileNotFoundError('Input File is not found.')

    try:
        # make instance for visio
        application = win32com.client.Dispatch(get_dispatch_format(in_extension[1:]))
        application.Visible = False
        document = application.Documents.Open(in_filename)

        # make pages of picture
        pages = get_pages(application, page_num=options.page)

        # define page_names
        if len(pages) == 1:
            page_names = [out_filename]
        else:   # len(pages) >= 2
            page_names = (out_filename_without_extension + str(page_cnt + 1) + out_extension
                    for page_cnt in range(len(pages)))

        # Export pages
        for page, page_name in zip(pages, page_names):
            page.Export(page_name)
    except com_error as err:
        raise IllegalImageFormatException('Output filename is not llegal for Image File.')
    finally:
        application.Quit()

    

if __name__ == '__main__':
    # define parser
    parser = OptionParser()
    parser.add_option(
            '-p', '--page',
            action='store',
            type='int',
            dest='page',
            help='transform only one page(set number of this page)'
        )
    (options, argv) = parser.parse_args()
    
    
    # if len(arguments) != 2, raise exception
    if len(argv) != 2:
        stderr.write('Enter Only input_filename and output_filename')
        exit()
    
    # define input_filename and output_filename
    in_filename = argv[0]
    out_filename = argv[1]

    try:
        export_img(in_filename, out_filename, options.page)
    except (FileNotFoundError, IllegalImageFormatException) as err:
                # expected exception
        print(str(err)) # print message
    except Exception as err:
        print(err.__traceback__)      # TraceBack Information
