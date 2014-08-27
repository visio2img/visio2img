import win32com.client
from win32com.client import constants
from os import path, chdir, getcwd

from sys import exit
from optparse import OptionParser


def get_dispatch_format(extension):
    if extension == 'vsd':
        return 'Visio.InvisibleApp'
    if extension == 'vsdx':
        pass    # pass


def get_pages(app, page_num=None):
    """
    app -> page
    if page_num is None, return all pages.
    if page_num is int object, return path_num-th page(fromm 1).
    """
    pages = app.ActiveDocument.Pages
    return [list(pages)[page_num - 1]] if page_num else pages

def export_img(in_filename, out_filename, page_num=None):
    """
    export as image format
    """
    # define filename without extension and extension variable
    in_filename_without_extension, in_extension = path.splitext(in_filename)
    out_filename_without_extension, out_extension = path.splitext(out_filename)

    # if file is not found, exit from program
    if not path.exists(in_filename):
        print('File Not Found')
        exit()

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
                    for page_cnt in range(0, len(pages)))

        # Export pages
        for page, page_name in zip(pages, page_names):
            page.Export(page_name)

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
        print('Enter Only input_filename and output_filename')
        exit()
    
    # define input_filename and output_filename
    in_filename = path.abspath(argv[0])
    out_filename = path.abspath(argv[1])
    
    export_img(in_filename, out_filename, options.page)
