#!/usr/bin/env python3

from sys import exit, stderr
try:
    import win32com.client
    from win32com.client import constants
except ImportError as err:
    stderr.write('win32com module not found')
    exit()

from pywintypes import com_error
from os import path, chdir, getcwd
from optparse import OptionParser

__all__ = ['export_img']

class IllegalImageFormatException(TypeError):
    """
    This exception means Exceptions for Illegal Image Format.
    """

class VisioNotFound(Exception):
    """
    This excetion means system has no visio program.
    """

def _get_dispatch_format(extension):
    return 'Visio.InvisibleApp' # vsd format


def _get_pages(app, page_num=None):
    """
    app -> page
    if page_num is None, return all pages.
    if page_num is int object, return path_num-th page(from 1).
    """
    pages = app.ActiveDocument.Pages
    return [list(pages)[page_num - 1]] if page_num else pages

def _check_format(visio_filename, gen_img_filename):
    visio_extension = path.splitext(visio_filename)[1]
    gen_img_extension = path.splitext(gen_img_filename)[1]
    if visio_extension not in ('.vsd'):
        err_str = (
                'Input filename is not llegal for visio file. \n' 
                'This program is suppert only vsd extension.'
                )
        raise IllegalImageFormatException(err_str)

    if gen_img_extension not in ('.gif', '.jpg', '.jpeg', '.png'):
                err_str = (
                'Output filename is not llegal for visio file. \n' 
                'This program is suppert gif, jpeg, png extension.'
                )
                raise IllegalImageFormatException(err_str)


def export_img(visio_filename, gen_img_filename, page_num=None, page_name=None):
    """
    export as image format
    """
    # to absolute path
    visio_filename = path.abspath(visio_filename)
    gen_img_filename = path.abspath(gen_img_filename)
    
    # define filename without extension and extension variable
    gen_img_filename_without_extension, gen_img_extension = path.splitext(gen_img_filename)

    _check_format(visio_filename, gen_img_filename)

    # if file is not found, exit from program
    if not path.exists(visio_filename):
        raise FileNotFoundError('Input File is not found.')

    gen_img_dir_name = gen_img_filename[:-len(path.basename(gen_img_filename))]
    if not path.isdir(gen_img_dir_name):
        raise FileNotFoundError('Directory of Output File is not found')

    try:
        # make instance for visio
        _, visio_extension = path.splitext(visio_filename)
        application = win32com.client.Dispatch(_get_dispatch_format(visio_extension[1:]))

        # case: system has no visio
        if application is None:
            raise VisioNotFoundException('System has no Visio.')

        application.Visible = False
        document = application.Documents.Open(visio_filename)

        # case: system has no visio
        if application is None:
            raise VisioNotFoundException('Visio Not Found in your system')

        # make pages of picture
        pages = _get_pages(application, page_num=options.page)

        ## filter of page names
        if page_name is not None:
            # dictionary of page to page.names
            page_dict = dict(zip(pages, pages.GetNames()))
            pages = list(filter(
                        lambda p: page_dict[page] ==  page_name,
                        pages))
        

        # define page_names
        if len(pages) == 1:
            page_names = [gen_img_filename]
        else:   # len(pages) >= 2
            _, visio_extension = path.splitext(visio_filename)
            page_names = (gen_img_filename_without_extension + str(page_cnt + 1) + gen_img_extension
                    for page_cnt in range(len(pages)))

        # Export pages
        for page, page_name in zip(pages, page_names):
            page.Export(page_name)
    except com_error as err:
        print(err)
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
    parser.add_option(
            '-n', '--name',
            action='store',
            type='string',
            dest='page_name',
            help='transform only same as setted name page'
        )
    (options, argv) = parser.parse_args()
    
    
    # if len(arguments) != 2, raise exception
    if len(argv) != 2:
        stderr.write('Enter Only input_filename and output_filename')
        exit()
    
    # define input_filename and output_filename
    visio_filename = argv[0]
    gen_img_filename = argv[1]

    try:
        export_img(visio_filename, gen_img_filename,
                           page_num=options.page,
                           page_name=options.page_name)
    except (FileNotFoundError, IllegalImageFormatException) as err:
                # expected exception
        stderr.write(str(err)) # print message
        """
    except Exception as err:
        print('Error')
        """
