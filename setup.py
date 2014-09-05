from setuptools import setup

setup(
    # program info
    name='visio2img',
    version='0.9.0',
    packages=['visio2img'],
    description=(
        'module or software for translation from visio format to'
        'other general image format.'),
    url='https://github.com/yassu/Visio2Img',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Console',
        'License :: Freeware',
        'License :: OSI Approved :: Apache Software License',
        'Intended Audience :: Developers'
        'Operating System :: Microsoft :: Windows',
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
