#!/usr/bin/env python3

import sys
from sys import stderr

from os import path
from optparse import OptionParser
from math import log

__all__ = ('export_img')

GEN_IMG_FORMATS = ('.gif', '.jpeg', '.jpg', '.png')


def is_pywin32_available():
    try:
        import win32com  # NOQA: import test
        return True
    except ImportError:
        return False


class FileNotFoundError(Exception):
    """
    exception represents the input file is not found
    """


class IllegalImageFormatException(TypeError):

    """
    This exception means Exceptions for Illegal Image Format.
    """


class UnsupportedFileError(Exception):
    """ exception represens the specified file is not supported """


class VisioNotFoundException(Exception):

    """
    This excetion means system has no visio program.
    """


def filter_pages(pages, pagenum, pagename):
    """ Choice pages using pagenum and pagename. """
    if pagenum:
        try:
            pages = list(pages)[pagenum - 1:pagenum]
        except IndexError:
            raise IndexError('Invalid page number: %d' % pagenum)

    if pagename:
        pages = [page for page in pages if pages.name == pagename]
        if pages == []:
            raise IndexError('Page not found: pagename=%s' % pagename)

    return pages


def _check_format(gen_img_filename):
    gen_img_extension = path.splitext(gen_img_filename)[1]
    if gen_img_extension not in GEN_IMG_FORMATS:
        errmsg = 'Unsupported image format: %s' % gen_img_filename
        raise IllegalImageFormatException(errmsg)


def export_img(visio_filename, gen_img_filename, pagenum=None, pagename=None):
    """
    export as image format
    If exported page, return True and else return False.
    """
    from pywintypes import com_error

    visio_pathname = path.abspath(visio_filename)
    gen_img_pathname = path.abspath(gen_img_filename)

    # define filename without extension and extension variable
    _check_format(gen_img_pathname)

    # if file is not found, exit from program
    if not path.exists(visio_pathname):
        raise FileNotFoundError('visio files not found: %s' % visio_filename)

    if not path.isdir(path.dirname(gen_img_pathname)):
        raise FileNotFoundError('Could not write image file: %s' % gen_img_filename)

    try:
        import win32com.client
        visioapp = win32com.client.Dispatch('Visio.InvisibleApp')
    except:
        raise VisioNotFoundException('Visio not found. visio2img requires Visio.')

    try:
        visioapp.Documents.Open(visio_pathname)
    except:
        raise UnsupportedFileError('Could not open file: %s' % visio_filename)

    try:
        pages = filter_pages(visioapp.ActiveDocument.Pages, pagenum, pagename)

        # define page_names
        if len(pages) == 1:
            page_names = [gen_img_pathname]
        else:   # len(pages) >= 2
            figure_length = int(log(len(pages), 10)) + 1
            gen_img_filename_without_extension, gen_img_extension = (
                 path.splitext(gen_img_pathname))
            page_names = (
                (gen_img_filename_without_extension +
                 ("{0:0>" + str(figure_length) + "}").format(page_cnt + 1) +
                 gen_img_extension
                 for page_cnt in range(len(pages))))

        # Export pages
        for page, page_name in zip(pages, page_names):
            page.Export(page_name)
        return True  # pages is not empty
    except com_error:
        raise IllegalImageFormatException(
            'Could not write image: %d' % gen_img_pathname)
    finally:
        visioapp.Quit()


def main(args=sys.argv[1:]):
    # define parser
    parser = OptionParser()
    parser.add_option(
        '-p', '--page',
        action='store',
        type='int',
        dest='pagenum',
        help='transform only one page(set number of this page)'
    )
    parser.add_option(
        '-n', '--name',
        action='store',
        type='string',
        dest='pagename',
        help='transform only same as setted name page'
    )
    (options, argv) = parser.parse_args(args)

    if (options.pagenum is not None) and (options.pagename is not None):
        stderr.write('--page and ---name options are conflicted')
        return -1

    # if len(arguments) != 2, raise exception
    if len(argv) != 2:
        parser.print_usage(stderr)
        return -1

    if not is_pywin32_available():
        stderr.write('win32com module not found')
        return -1

    # define input_filename and output_filename
    visio_filename = argv[0]
    gen_img_filename = argv[1]

    try:
        export_img(visio_filename, gen_img_filename,
                   options.pagenum, options.pagename)

        return 0
    except (FileNotFoundError, VisioNotFoundException, IllegalImageFormatException, IndexError) as err:
        # expected exception
        stderr.write(str(err))  # print message
        return -1
