# -*- coding: utf-8 -*-

import os
import sys
import unittest
from shutil import rmtree
from tempfile import mkdtemp
from collections import namedtuple

from visio2img.visio2img import (
    is_pywin32_available,
    filter_pages,
    export_img,
    main,
)

if sys.version_info > (3, 0):
    from unittest.mock import patch
else:
    from mock import patch

EXAMPLE_DIR = os.path.join(os.path.dirname(__file__), 'examples')

VISIO_AVAILABLE = False
if is_pywin32_available():
    import win32com.client
    try:
        app = win32com.client.Dispatch('Visio.InvisibleApp')
        app.Quit()
        VISIO_AVAILABLE = True
    except Exception:
        pass


class TestVisio2img(unittest.TestCase):
    def test_is_pywin32_available(self):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules.pop('win32com', None)  # unload forcely

            self.assertFalse(is_pywin32_available())

            sys.modules['win32com'] = True
            self.assertTrue(is_pywin32_available())
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload forcely

    def test_filter_pages_by_default(self):
        pages = range(10)
        self.assertEqual(pages, filter_pages(pages, None, None))

    def test_filter_pages_by_pagenum(self):
        pages = range(10)
        self.assertEqual([4], filter_pages(pages, 5, None))

        with self.assertRaises(IndexError):
            filter_pages(pages, 100, None)

    def test_filter_pages_by_pagename(self):
        Page = namedtuple('Page', 'num name')
        pages = [Page(n, "Page %d" % n) for n in range(10)]
        self.assertEqual([(5, 'Page 5')], filter_pages(pages, None, 'Page 5'))

        with self.assertRaises(IndexError):
            filter_pages(pages, None, 'unknown')

    @patch("sys.stderr")
    @patch("visio2img.visio2img.export_img")
    def test_parse_option(self, export_img, _):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules['win32com'] = True

            # no arguments, win32com available
            with self.assertRaises(SystemExit):
                args = []
                main(args)
            self.assertEqual(0, export_img.call_count)

            # one argument, win32com available
            with self.assertRaises(SystemExit):
                args = ['input.vsd']
                main(args)
            self.assertEqual(0, export_img.call_count)

            # two arguments, win32com available
            args = ['input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(0, ret)
            self.assertEqual(1, export_img.call_count)
            export_img.assert_called_with('input.vsd', 'output.png',
                                          None, None)

            # three arguments, win32com available
            with self.assertRaises(SystemExit):
                args = ['input.vsd', 'output.png', 'other_args']
                main(args)
            self.assertEqual(1, export_img.call_count)

            # two arguments, --page option, win32com available
            args = ['-p', '3', 'input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(0, ret)
            self.assertEqual(2, export_img.call_count)
            export_img.assert_called_with('input.vsd', 'output.png',
                                          3, None)

            # two arguments, --name option, win32com available
            args = ['-n', 'sheet1', 'input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(0, ret)
            self.assertEqual(3, export_img.call_count)
            export_img.assert_called_with('input.vsd', 'output.png',
                                          None, 'sheet1')

            # two arguments, --page and --name option, win32com available
            with self.assertRaises(SystemExit):
                args = ['-p', '3', '-n', 'sheet1', 'input.vsd', 'output.png']
                main(args)
            self.assertEqual(3, export_img.call_count)

            # two arguments, win32com unavailable
            sys.modules.pop('win32com', None)  # unload forcely
            args = ['input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(-1, ret)
            self.assertEqual(3, export_img.call_count)
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload forcely

    @patch("sys.stderr")
    @patch("visio2img.visio2img.export_img")
    def test_check_image_formats(self, export_img, _):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules['win32com'] = True

            # .png
            args = ['input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(0, ret)

            # .gif
            args = ['input.vsd', 'output.gif']
            ret = main(args)
            self.assertEqual(0, ret)

            # .jpg
            args = ['input.vsd', 'output.jpg']
            ret = main(args)
            self.assertEqual(0, ret)

            # .pdf
            with self.assertRaises(SystemExit):
                args = ['input.vsd', 'output.pdf']
                main(args)

            # .PNG (capital)
            args = ['input.vsd', 'output.PNG']
            ret = main(args)
            self.assertEqual(0, ret)

            # no extension
            with self.assertRaises(SystemExit):
                args = ['input.vsd', 'output_without_ext']
                main(args)
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload forcely
            sys.modules.pop('win32com.client', None)  # unload forcely

    @patch("sys.stderr")
    @patch("visio2img.visio2img.export_img")
    def test_main_if_export_img_raises_error(self, export_img, _):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules['win32com'] = True
            args = ['input.vsd', 'output.png']

            # case of IOError
            export_img.side_effect = IOError
            ret = main(args)
            self.assertEqual(-1, ret)

            # case of OSError
            export_img.side_effect = OSError
            ret = main(args)
            self.assertEqual(-1, ret)

            # case of IndexError
            export_img.side_effect = IndexError
            ret = main(args)
            self.assertEqual(-1, ret)

            # case of other exception (does not handle it)
            export_img.side_effect = Exception
            with self.assertRaises(Exception):
                main(args)
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload forcely
            sys.modules.pop('win32com.client', None)  # unload forcely

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_singlepage_to_png(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'singlepage.vsdx'),
                       os.path.join(tmpdir, 'output.png'), None, None)
            self.assertEqual(['output.png'], os.listdir(tmpdir))

            expected = os.path.join(EXAMPLE_DIR, 'singlepage', 'output.png')
            actual = os.path.join(tmpdir, 'output.png')
            self.assertEqual(open(expected, 'rb').read(),
                             open(actual, 'rb').read())
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_singlepage_to_jpg(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'singlepage.vsdx'),
                       os.path.join(tmpdir, 'output.jpg'), None, None)
            self.assertEqual(['output.jpg'], os.listdir(tmpdir))

            expected = os.path.join(EXAMPLE_DIR, 'singlepage', 'output.jpg')
            actual = os.path.join(tmpdir, 'output.jpg')
            self.assertEqual(open(expected, 'rb').read(),
                             open(actual, 'rb').read())
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_singlepage_to_gif(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'singlepage.vsdx'),
                       os.path.join(tmpdir, 'output.gif'), None, None)
            self.assertEqual(['output.gif'], os.listdir(tmpdir))

            expected = os.path.join(EXAMPLE_DIR, 'singlepage', 'output.gif')
            actual = os.path.join(tmpdir, 'output.gif')
            self.assertEqual(open(expected, 'rb').read(),
                             open(actual, 'rb').read())
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_multipages1(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'multipages.vsdx'),
                       os.path.join(tmpdir, 'output.png'), None, None)
            self.assertEqual(['output1.png', 'output2.png'],
                             os.listdir(tmpdir))
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_multipages2(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'multipages2.vsdx'),
                       os.path.join(tmpdir, 'output.png'), None, None)

            expected = ['output01.png', 'output02.png', 'output03.png',
                        'output04.png', 'output05.png', 'output06.png',
                        'output07.png', 'output08.png', 'output09.png',
                        'output10.png']
            self.assertEqual(expected, os.listdir(tmpdir))
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_multipages_with_pagenum(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'multipages.vsdx'),
                       os.path.join(tmpdir, 'output.png'), 2, None)
            self.assertEqual(['output.png'], os.listdir(tmpdir))

            expected = os.path.join(EXAMPLE_DIR, 'multipages', 'output2.png')
            actual = os.path.join(tmpdir, 'output.png')
            self.assertEqual(open(expected, 'rb').read(),
                             open(actual, 'rb').read())

        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_multipages_with_pagename(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'multipages.vsdx'),
                       os.path.join(tmpdir, 'output.png'), None, u"ページ - 2")
            self.assertEqual(['output.png'], os.listdir(tmpdir))

            expected = os.path.join(EXAMPLE_DIR, 'multipages', 'output2.png')
            actual = os.path.join(tmpdir, 'output.png')
            self.assertEqual(open(expected, 'rb').read(),
                             open(actual, 'rb').read())

        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_from_vsd(self):
        try:
            tmpdir = mkdtemp()
            export_img(os.path.join(EXAMPLE_DIR, 'multipages.vsd'),
                       os.path.join(tmpdir, 'output.png'), None, None)
            self.assertEqual(['output1.png', 'output2.png'],
                             os.listdir(tmpdir))
        finally:
            rmtree(tmpdir)

    def test_export_img_visio_file_not_found(self):
        with self.assertRaises(IOError):
            export_img('/path/to/notexist.vsd', '/path/to/output.png',
                       None, None)

    def test_export_img_output_dir_not_found(self):
        with self.assertRaises(IOError):
            export_img(os.path.join(EXAMPLE_DIR, 'singlepage.vsdx'),
                       '/path/to/output.png', None, None)

    @patch('win32com.client')
    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_if_visio_not_found(self, win32com_client):
        from pywintypes import com_error
        win32com_client.Dispatch.side_effect = com_error

        try:
            tmpdir = mkdtemp()
            with self.assertRaises(OSError):
                export_img(os.path.join(EXAMPLE_DIR, 'singlepage.vsdx'),
                           os.path.join(tmpdir, 'output.png'), None, None)
        finally:
            rmtree(tmpdir)

    @unittest.skipIf(VISIO_AVAILABLE is False, "Visio not found")
    def test_export_img_with_non_visio_file(self):
        try:
            tmpdir = mkdtemp()
            with self.assertRaises(IOError):
                export_img(__file__,
                           os.path.join(tmpdir, 'output.png'), None, None)
        finally:
            rmtree(tmpdir)
