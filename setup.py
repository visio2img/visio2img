import sys
from setuptools import setup

classifiers = [
    'Development Status :: 4 - Beta',
    'Environment :: Console',
    'Environment :: Win32 (MS Windows)',
    'Intended Audience :: System Administrators',
    'License :: OSI Approved :: Apache Software License',
    'Operating System :: Microsoft :: Windows',
    'Programming Language :: Python',
    'Programming Language :: Python :: 2',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.3',
    'Programming Language :: Python :: 3.4',
    'Topic :: Documentation',
    'Topic :: Multimedia :: Graphics :: Graphics Conversion',
    'Topic :: Office/Business :: Office Suites',
    'Topic :: Software Development :: Libraries :: Python Modules',
    'Topic :: Utilities',
]

if sys.version_info > (3, 0):
    test_requires = []
else:
    test_requires = ['mock']

setup(
    name='visio2img',
    version='1.0.0',
    description='MS-Visio file (.vsd, .vsdx) to images converter',
    long_description=(
        'If you use this program as command in terminal, '
        'this program provides visio2img.py command.\n'
        'If you use this program as module of python, '
        'this module provides visio2img.visio2img.export_img function.\n'
        'Requirements of this program is '
        'Visio application and win32com module.\n'
        'This program is for only python3.'
    ),
    author='Yassu',
    author_email='yassumath@gmail.com',
    maintainer='Takeshi KOMIYA',
    maintainer_email='i.tkomiya@gmail.com',
    url='https://github.com/visio2img/visio2img',
    classifiers=classifiers,
    packages=['visio2img'],
    tests_require=test_requires,
    entry_points="""
       [console_scripts]
       visio2img = visio2img.visio2img:main
    """
)
