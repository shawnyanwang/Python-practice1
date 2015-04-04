__author__ = 'shawn.wang'

from distutils.core import setup
import py2exe

setup(
    windows=[{'script': 'Interface_web_scrape.py'}],
    options={
        'py2exe':
        {
            'includes': ['lxml.etree', 'lxml._elementpath', 'gzip'],
        }
    }
)