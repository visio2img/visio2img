import win32com.client
from win32com.client import constants
from os import path, chdir, getcwd
from sys import argv, exit


def get_dispatch_format(extension):
    if extension == 'vsd':
        return 'Visio.Application'
    if extension == 'vsdx':
        pass    # What?
    

if __name__ == '__main__':
    # if len(arguments) != 2, raise exception
    if len(argv) != 3:
        print('Enter Only input_filename and output_filename')
    
    # define input_filename and output_filename
    in_filename = path.abspath(argv[1])
    out_filename = path.abspath(argv[2])

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
        pages = application.ActiveDocument.Pages

        # define page_names
        if len(pages) == 1:
            page_names = [out_filename]
        else:   # len(pages) >= 2
            page_names = (out_filename_without_extension + str(page_cnt + 1) + '.png'
                    for page_cnt in range(0, len(pages)))

        # Export pages
        for page, page_name in zip(pages, page_names):
            page.Export(page_name)

    finally:
        application.Quit()
