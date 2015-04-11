import sys
from os import path
import json
from IPython.nbformat import read, write
from common import NOTEBOOK_DIR
from diff_editor import iter_notebook_paragraphs, diff_html

name = sys.argv[1]

try:
    is_apply = sys.argv[2] == "apply"
except IndexError:
    is_apply = False

with open(path.join(NOTEBOOK_DIR, name, "changes.json"), "rb") as f:
    changes = json.load(f)

def iter_changes(changes):
    for key, ftext in changes.iteritems():
        nb_name, cell_index = key.split(" - ")
        cell_index = int(cell_index)
        if cell_index >= 0:
            yield key, nb_name, cell_index, ftext

if not is_apply:
    texts = {"{} - {}".format(name, index):text for name, index, text in
                  iter_notebook_paragraphs(path.join(NOTEBOOK_DIR, name))}

    with open(path.join(NOTEBOOK_DIR, name, "changes2.html"), "wb") as f:
        f.write("<html><body>")
        for key, nb_name, cell_index, ftext in iter_changes(changes):
            otext = texts[key]
            f.write("<h2>{}</h2>".format(key))
            f.write("<p>{}</p>".format(diff_html(otext, ftext).encode("utf8")))
        f.write("</body></html>")
else:
    notebooks = {}
    for key, nb_name, cell_index, ftext in iter_changes(changes):
        print key, nb_name, cell_index
        if nb_name not in notebooks:
            with open(path.join(NOTEBOOK_DIR, name, nb_name), "rb") as f:
                notebooks[nb_name] = read(f, 4)
        notebook = notebooks[nb_name]
        cell = notebook.cells[cell_index]
        cell["source"] = ftext
    for nb_name, notebook in notebooks.iteritems():
        write(notebook, path.join(NOTEBOOK_DIR, name, nb_name))
