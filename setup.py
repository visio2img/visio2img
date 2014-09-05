from setuptools import setup

setup(
    # program info
    name='visio2img',
    version='1.0.0',
    packages=['visio2img'],
    description=(
        'module or software for translation from visio format to'
        'other general image format.'),
    long_description=(
        'If you use this program as command in terminal, ' 
            'this program provides visio2img.py command.\n'
        'If you use this program as module of python, '
            'this module provides visio2img.visio2img.export_img function.\n'
        'Requirements of this program is '
            'Visio application and win32com module.\n'
            'This program is for only python3.'
            ),
    url='https://github.com/yassu/Visio2Img',
    classifiers=[
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Development Status :: 4 - Beta',
        'Environment :: Console',
        'Topic :: Software Development :: Libraries :: Application Frameworks',
            # for my sphinxcontrib-visio
        'License :: Freeware',
        'License :: OSI Approved :: Apache Software License',
        'Intended Audience :: Developers'
        'Operating System :: Microsoft :: Windows',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Software Development :: Embedded Systems',
        'Topic :: Office/Business'
    ],
    license=(
        'Released Under the Apache license\n'
        'https://github.com/yassu/Visio2Img\n'
    ),
    scripts=['visio2img/visio2img.py'],

    # author info
    author='Yassu',
    author_email='yassumath@gmail.com',
)
