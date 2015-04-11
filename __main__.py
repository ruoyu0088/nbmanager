import argparse
from os import path, walk
from fnmatch import fnmatch
from nbmanager.docxformater import process_docx
from nbmanager.mdcreater import convert_md_to_docx, convert_folder_to_md
from .common import BUILD_ROOT_FOLDER

parser = argparse.ArgumentParser(description='convert notebooks to docx')
parser.add_argument('folder', type=str, help='folder')
parser.add_argument('-p', '--pattern', type=str, help='pattern', default="")
parser.add_argument('-s', '--step', default="all",
                   help='step to execute', choices=["all", "md", "docx", "pdocx"])
args = parser.parse_args()

if args.pattern == "":
    args.pattern = args.folder.split("-")[1] + "-*.ipynb"

if args.step == "md" or args.step == "all":
    convert_folder_to_md(args.folder, args.pattern)

if args.step == "docx" or args.step == "all":
    md_file = path.abspath(path.join(BUILD_ROOT_FOLDER, r"build_{0}\{0}.md".format(args.folder)))
    convert_md_to_docx(md_file)

if args.step == "pdocx" or args.step == "all":
    process_docx(args.folder)