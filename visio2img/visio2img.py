# -*- coding: utf-8 -*-
#  Copyright 2014 Yassu
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import os
import sys
from optparse import OptionParser
from math import log

__all__ = ('export_img')


def is_pywin32_available():
    """ Tests pywin32 is installed """
    try:
        import win32com  # NOQA: import test
        return True
    except ImportError:
        return False


def filter_pages(pages, pagenum, pagename):
    """ Choices pages by pagenum and pagename """
    if pagenum:
        try:
            pages = [list(pages)[pagenum - 1]]
        except IndexError:
            raise IndexError('Invalid page number: %d' % pagenum)

    if pagename:
        pages = [page for page in pages if page.name == pagename]
        if pages == []:
            raise IndexError('Page not found: pagename=%s' % pagename)

    return pages


def export_img(visio_filename, gen_img_filename, pagenum=None, pagename=None):
    """ Exports images from visio file """
    from pywintypes import com_error

    # visio requires absolute path
    visio_pathname = os.path.abspath(visio_filename)
    gen_img_pathname = os.path.abspath(gen_img_filename)

    if not os.path.exists(visio_pathname):
        raise IOError('No such visio file: %s', visio_filename)

    if not os.path.isdir(os.path.dirname(gen_img_pathname)):
        msg = 'Could not write image file: %s' % gen_img_filename
        raise IOError(msg)

    try:
        import win32com.client
        visioapp = win32com.client.Dispatch('Visio.InvisibleApp')
    except:
        msg = 'Visio not found. visio2img requires Visio.'
        raise OSError(msg)

    try:
        visioapp.Documents.Open(visio_pathname)
    except:
        msg = 'Could not open file (already opend by other process?): %s'
        raise IOError(msg % visio_filename)

    try:
        pages = filter_pages(visioapp.ActiveDocument.Pages, pagenum, pagename)

        if len(pages) == 1:
            pages[0].Export(gen_img_pathname)
        else:
            digits = int(log(len(pages), 10)) + 1
            basename, ext = os.path.splitext(gen_img_pathname)
            filename_format = "%s%%0%dd%s" % (basename, digits, ext)

            for i, page in enumerate(pages):
                img_filename = filename_format % (i + 1)
                page.Export(img_filename)
    except com_error:
        raise IOError('Could not write image: %d' % gen_img_pathname)
    finally:
        visioapp.Quit()


def parse_options(args):
    """ Parses command line options """
    parser = OptionParser()
    parser.add_option('-p', '--page', action='store',
                      type='int', dest='pagenum',
                      help='pick a page by page number')
    parser.add_option('-n', '--name', action='store',
                      type='string', dest='pagename',
                      help='pick a page by page name')
    options, argv = parser.parse_args(args)

    if options.pagenum and options.pagename:
        parser.error('options --page and --name are mutually exclusive')

    if len(argv) != 2:
        parser.print_usage(sys.stderr)
        parser.exit()

    output_ext = os.path.splitext(argv[1])[1].lower()
    if output_ext not in ('.gif', '.jpg', '.png'):
        parser.error('Unsupported image format: %s' % argv[1])

    return options, argv


def main(args=sys.argv[1:]):
    """ main funcion of visio2img """
    if not is_pywin32_available():
        sys.stderr.write('win32com module not found')
        return -1

    try:
        options, argv = parse_options(args)
        export_img(argv[0], argv[1], options.pagenum, options.pagename)
        return 0
    except (IOError, OSError, IndexError) as err:
        sys.stderr.write("error: %s" % err)
        return -1
