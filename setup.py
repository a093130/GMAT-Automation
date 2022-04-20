#!python
# # -*- coding: utf-8 -*-
"""
Created on Sun Apr 29 15:08:37 2019

@author: Colin
@author_email: colinhelms@outlook.com
"""

from distutils.core import setup
setup(name='gmatautomation',
      version='0.2a',
      description='GMAT model generation, batch execution, and data reduction.',
      long_description='file: README.txt',
      author='Colin Helms',
      author_email='colinhelms@outlook.com',
      url='https://www.FreelanceRocketScience.com/downloads',
      packages=['src', 'src.modelcontrol', 'src.reportgen', 'src.modelgen'],
      #package_data={'modelcontrol' : ['']},
      data_files= [('', 'docs.GMATAutomation_SoftwareUserManual.docx')],
      classifiers=[
              'Development Status :: 1 - Alpha',
              'Environment :: Console',
              'Intended Audience :: End Users',
              'Operating System :: Microsoft :: Windows7',
              'Programming Language :: Python :: 3.4',
              'License :: End User License Agreement',
              'Topic :: Copyright :: Copyright Freelance Rocket Science, 2019']
      )