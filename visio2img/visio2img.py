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
from math import log
from optparse import OptionParser


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


class VisioFile(object):
    @classmethod
    def Open(cls, filename):
        obj = cls()
        obj.open(filename)
        return obj

    def __init__(self):
        self.app = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
        return False

    def open(self, filename):
        assert self.app is None

        visio_pathname = os.path.abspath(filename)  # visio requires abspath
        if not os.path.exists(visio_pathname):
            raise IOError('No such visio file: %s', filename)

        try:
            import win32com.client
            self.app = win32com.client.Dispatch('Visio.InvisibleApp')
        except:
            msg = 'Visio not found. visio2img requires Visio.'
            raise OSError(msg)

        try:
            if hasattr(self.app.Documents, "OpenEx"):
                # Visio >= 4.5 supports OpenEx
                # visOpenCopy + visOpenRO allows opening documents even
                # if they're open in another visio instance...
                visOpenCopy = 0x1
                visOpenRO = 0x2
                open_flags = visOpenCopy | visOpenRO
                self.app.Documents.OpenEx(visio_pathname, open_flags)
            else:
                self.app.Documents.Open(visio_pathname)
        except:
            self.close()
            msg = 'Could not open file (already opend by other process?): %s'
            raise IOError(msg % filename)

    def close(self):
        if self.app:
            self.app.Quit()
            self.app = None

    @property
    def pages(self):
        if self.app:
            return self.app.ActiveDocument.Pages
        else:
            return []


def export_img(visio_filename, image_filename, pagenum=None, pagename=None):
    """ Exports images from visio file """
    # visio requires absolute path
    image_pathname = os.path.abspath(image_filename)

    if not os.path.isdir(os.path.dirname(image_pathname)):
        msg = 'Could not write image file: %s' % image_filename
        raise IOError(msg)

    with VisioFile.Open(visio_filename) as visio:
        pages = filter_pages(visio.pages, pagenum, pagename)
        try:
            if len(pages) == 1:
                pages[0].Export(image_pathname)
            else:
                digits = int(log(len(pages), 10)) + 1
                basename, ext = os.path.splitext(image_pathname)
                filename_format = "%s%%0%dd%s" % (basename, digits, ext)

                for i, page in enumerate(pages):
                    filename = filename_format % (i + 1)
                    page.Export(filename)
        except:
            raise IOError('Could not write image: %s' % image_pathname)


def parse_options(args):
    """ Parses command line options """
    usage = 'usage: %prog [options] visio_filename image_filename'
    parser = OptionParser(usage=usage)
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
