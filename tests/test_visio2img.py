import sys
import unittest
from unittest.mock import patch
from collections import namedtuple
from visio2img.visio2img import (
    is_pywin32_available,
    filter_pages,
    _check_format,
    main,
    FileNotFoundError,
    VisioNotFoundException,
    IllegalImageFormatException,
)


class TestVisio2img(unittest.TestCase):
    def test_is_pywin32_available(self):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules.pop('win32com', None)  # unload win32com forcely

            self.assertFalse(is_pywin32_available())

            sys.modules['win32com'] = True
            self.assertTrue(is_pywin32_available())
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload win32com forcely

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

    def test_check_format(self):
        self.assertIsNone(_check_format('image.gif'))
        self.assertIsNone(_check_format('image.jpg'))
        self.assertIsNone(_check_format('image.jpeg'))
        self.assertIsNone(_check_format('image.png'))

        with self.assertRaises(IllegalImageFormatException):
            _check_format('image.pdf')

        with self.assertRaises(IllegalImageFormatException):
            _check_format('filename_without_ext')

    @patch("visio2img.visio2img.stderr")
    @patch("visio2img.visio2img.export_img")
    def test_parse_option(self, export_img, _):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules['win32com'] = True

            # no arguments, win32com available
            args = []
            ret = main(args)
            self.assertEqual(-1, ret)
            self.assertEqual(0, export_img.call_count)

            # one argument, win32com available
            args = ['input.vsd']
            ret = main(args)
            self.assertEqual(-1, ret)
            self.assertEqual(0, export_img.call_count)

            # two arguments, win32com available
            args = ['input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(0, ret)
            self.assertEqual(1, export_img.call_count)
            export_img.assert_called_with('input.vsd', 'output.png',
                                          None, None)

            # three arguments, win32com available
            args = ['input.vsd', 'output.png', 'other_args']
            ret = main(args)
            self.assertEqual(-1, ret)
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
            args = ['-p', '3', '-n', 'sheet1', 'input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(-1, ret)
            self.assertEqual(3, export_img.call_count)

            # two arguments, win32com unavailable
            sys.modules.pop('win32com', None)  # unload win32com forcely
            args = ['input.vsd', 'output.png']
            ret = main(args)
            self.assertEqual(-1, ret)
            self.assertEqual(3, export_img.call_count)
        finally:
            sys.path = loadpath  # write back library loading paths
            sys.modules.pop('win32com', None)  # unload win32com forcely

    @patch("visio2img.visio2img.stderr")
    @patch("visio2img.visio2img.export_img")
    def test_main_if_export_img_raises_error(self, export_img, _):
        try:
            loadpath, sys.path = sys.path, []  # disable to load all modules
            sys.modules['win32com'] = True
            args = ['input.vsd', 'output.png']

            # case of FileNotFoundError
            export_img.side_effect = FileNotFoundError
            ret = main(args)
            self.assertEqual(-1, ret)

            # case of VisioNotFoundException
            export_img.side_effect = VisioNotFoundException
            ret = main(args)
            self.assertEqual(-1, ret)

            # case of IllegalImageFormatException
            export_img.side_effect = IllegalImageFormatException
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
            sys.modules.pop('win32com', None)  # unload win32com forcely
