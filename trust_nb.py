import subprocess
import sys
from os import path
from glob import glob
from common import iter_notebooks

for notebook_filename in iter_notebooks():
    print notebook_filename
    subprocess.call(["ipython", "trust", notebook_filename, "--profile", "scipybook2"])