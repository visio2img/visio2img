=========
visio2img
=========

`visio2img` is a image converter. It converts from MS-Visio file (.vsd, .vsdx) to images.

.. image:: https://travis-ci.org/visio2img/visio2img.svg?branch=master
   :target: https://travis-ci.org/visio2img/visio2img
   :alt: Build Status

.. image:: https://coveralls.io/repos/visio2img/visio2img/badge.png?branch=master
   :target: https://coveralls.io/r/visio2img/visio2img?branch=master
   :alt: Coverage

.. image:: https://pypip.in/v/visio2img/badge.png
   :target: https://pypi.python.org/pypi/visio2img/
   :alt: Latest PyPI version

.. image:: https://pypip.in/d/visio2img/badge.png
   :target: https://pypi.python.org/pypi/visio2img/
   :alt: Number of PyPI downloads

Requirements
=============

* Python 2.7, 3.3 and later
* pywin32_
* Microsoft Visio

.. _pywin32: http://sourceforge.net/projects/pywin32/files/pywin32/

Setup
=====

1. Install pywin32_ manually
2. Install `visio2img` package::

     $ pip install visio2img

And then, `visio2img` command is available on your environment.

Usage
======

Execute `visio2img` command::

   $ visio2img [visio_filename.vsdx] [image_filename.png]

If your visio file has multiple pages, `visio2img` command generates image files for each page.

page option
------------

`-p` (`--page`) option choices a page by page number.

For example, this command-line picks up second page of visio file::

   visio2img.py -p 2 visio_filename.vsdx output.png

name option
------------

`-n` (`--name`) option choices a page by page name.

For example, this command-line picks up a page named "circle"::

   visio2img.py -n "circle" visio_filename.vsdx output.png

Author
=======

Yassu <mathyassu@gmail.com>

Maintainer
===========

Takeshi KOMIYA <i.tkomiya@gmail.com>

License
========
Apache License 2.0
