
import os
import py_compile
import shutil
import csv
from random import random
from os.path import isfile, join, basename





#  '/home/anon/Documentos/pagBank/'

path_origem =  'fomm\\sptx.json'
path_destino = 'C:\\ProgramData\\Dig-bin\\'

try:
    os.mkdir(path_destino)
    print('.....')
    
except Exception as e:
    pass

try:
    shutil.move(path_origem, join(path_destino, basename(path_origem)))
except:
    pass

py_compile.compile('fomm\\DigbcPaulista.py')



exit(69)