from setuptools import setup

import unittest
def test_suite():
    test_loader = unittest.TestLoader()
    test_suite = test_loader.discover('merge_xlsx.tests', pattern='test_*.py')
    return test_suite

setup(name='merge_xlsx',
      version='0.1',
      description='Excel templates',
      author=u'Marcos S\xe1nchez Provencio',
      author_email='marcos@meteogrid.com',
      url='https://www.meteogrid.com/',
      packages=['merge_xlsx'],
      install_requires=[
                        'path.py>=7.7.1',
                        'openpyxl>=2.4',
                        'unittest2>=0.5.1'
                        ],
      test_suite='setup.test_suite',
     )
