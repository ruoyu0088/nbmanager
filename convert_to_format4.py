from IPython import nbformat
from .common import iter_notebooks

for notebook_filename in iter_notebooks():
    print "converting", notebook_filename
    book = nbformat.read(notebook_filename, 4)
    nbformat.write(book, notebook_filename)