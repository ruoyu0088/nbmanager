import os
from os import path
import re
from scpy2.utils.program_finder import InkscapePath

INKSCAPE_PATH = unicode(InkscapePath)
NOTEBOOK_DIR = os.environ["BOOK_ROOT"]
FOLDER = path.dirname(__file__)
BUILD_ROOT_FOLDER = path.join(NOTEBOOK_DIR, path.pardir, "build")

if not path.exists(BUILD_ROOT_FOLDER):
    os.mkdir(BUILD_ROOT_FOLDER)


def iter_chapters():
    for folder, subfolder, files in os.walk(NOTEBOOK_DIR):
        name = path.basename(folder)
        if re.match(r"\d+-\w+", name):
            yield folder


def iter_notebooks(chapters=None):
    if chapters is None:
        chapters = iter_chapters()
    for folder in chapters:
            for fn in os.listdir(folder):
                if fn.startswith("_"):
                    continue
                if re.match(r"\w+-\d\d\d-.+?\.ipynb", fn):
                    fn = path.join(folder, fn)
                    yield path.abspath(fn)
