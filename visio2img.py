import win32com.client
from win32com.client import constants
from os import path, chdir, getcwd
from sys import argv


try:
    # set filename
    filename = r'C:\Users\yuki_yasuda\Proj\test2.vsd'
    directory = '\\'.join(filename.split('\\')[:-1]) + '\\'
    print(directory)
    #filename = path.abspath(argv[0])

    print('file exists? {}'.format(path.exists(filename)))

    # make visio format instance
    #win32com.client.gencache.EnsureDispatch("Visio.Application")
    application = win32com.client.Dispatch("Visio.Application")
    application.Visible = False
    document = application.Documents.Open(filename)

    pages = application.ActiveDocument.Pages

    for page in pages:
        pagename = '%s%s.png' % (directory, page)
        print(pagename)
        page.Export(pagename)
finally:
    application.Quit()
