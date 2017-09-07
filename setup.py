import sys
from setuptools import setup

import unittest
def test_suite():
    test_loader = unittest.TestLoader()
    test_suite = test_loader.discover('merge_xlsx.tests', pattern='test_*.py')
    return test_suite

kw = {}
if sys.version_info < (2,7):
    kw.update(install_requires=[
            'path==2.2',
            'jdcal',
                        ],
      tests_require=[
                        'unittest2>=0.5.1'
                        ],
    )
else:
    kw.update(
              install_requires=[
                                'path.py',
        ],
              tests_require=[
                        'openpyxl>=2.4',
                        'unittest2>=0.5.1'
                        ],
    )
          

setup(name='merge_xlsx',
      version='0.1',
      description='Excel templates',
      author=u'Marcos S\xe1nchez Provencio',
      author_email='marcos@meteogrid.com',
      url='https://www.meteogrid.com/',
      packages=['merge_xlsx'],
      test_suite='setup.test_suite',
      **kw
      )
