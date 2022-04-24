#!python
# # -*- coding: utf-8 -*-
"""
Created on Sun Apr 24 04:22:37 2022

@author: Colin Helms
"""

from distutils.core import setup
setup(name='gmatautomation',
      version='0.2a',
      description='GMAT model generation, batch execution, and data reduction.',
      long_description='file: README.md',
      author='Colin Helms',
      author_email='colinhelms@outlook.com',
      url='https://github.com/a093130/GMAT-Automation',
      packages=['', 'gmatautomation', 'gmatautomation.modelcontrol', 'gmatautomation.modelgen', 'gmatautomation.reportgen'],
      classifiers=[
              'Development Status :: 1 - Alpha',
              'Environment :: Console',
              'Operating System :: Microsoft :: Windows 10',
              'Programming Language :: Python :: 3.7',
              'License :: GNU GENERAL PUBLIC LICENSE V3',
              'Topic :: Copyright :: Copyright Freelance Rocket Science, 2022'
              ],
        )