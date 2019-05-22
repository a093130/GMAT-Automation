# -*- coding: utf-8 -*-
"""
Created on Sun Apr 29 15:08:37 2019

@author: Colin
"""

from distutils.core import setup
setup(name='gmatautomod',
      version='0.1a',
      author='Colin Helms',
      author_email='colinhelms@outlook.com',
      url='https://www.FreelanceRocketScience.com/downloads',
      packages=['modelcontrol', 'modelgen'],
      package_data={'modelgen' : ['Vehicle Optimization Tables JS&R_R8.xlsx']},
      data_files= [('', 'GMAT Automation Software User Manual.docx')],
      scripts=['modelcontrol/gmat_batcher', 'modelcontrol/reduce_report', 'modelgen/modelget'], 
      classifiers=[
              'Development Status :: 1 - Alpha',
              'Environment :: Console',
              'Intended Audience :: End Users',
              'Operating System :: Microsoft :: Windows7',
              'Programming Language :: Python :: 3.5',
              'License :: End User License Agreement',
              'Topic :: Copyright :: Copyright Freelance Rocket Science, 2019']
      )