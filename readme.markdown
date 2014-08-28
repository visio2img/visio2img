## Introduction
 This is program for translation vsd to other image format.

## Usage

 If you use this program as module, this program provides  `visio2img.visio2img.export_img(in_filename, out_filename)` function.
 Assume that format of `in_filename` is vsd and format of `out_filename` is general image file name(format is judged from extension).
 
 For using command, Usage of this program is
```
visio2img.py {input_filename} {output_filename} [options ...] 
```
.
If number of page of input file is one, this program make a picture file named output filename.
If number of page of input file is more than one, this program make picture files named output filename1, output filename2, e.t.c.
For example, following command make files called named1.jpg, named2.jpg, named3.jpg:

```
visio2img.py 3pages.vsd named.jpg
```
, where 3pages.vsd is a visio file which have three pages.

## Requirements

* python3 \
	This program is for python3.
* win32com \
	Because of use for Visio.
	
## Options

 Now, The optional number has only one.
This is `-p` or `--page` option.
This is used for appointing a page number of visio file. 
For example

```
visio2img.py in.vsd out.jpg -p 1
```

for translate only first page.

## Known Problems

* A white screen appears for an instant.
* Only vsd: This program now support only vsd, little old format.