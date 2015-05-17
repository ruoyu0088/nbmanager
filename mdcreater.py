# -*- coding: utf-8 -*-
import os
from os import path
import re
import subprocess
import shutil
from glob import glob
from IPython.nbconvert.exporters.markdown import MarkdownExporter
from IPython.nbconvert.writers import FilesWriter
import pandas as pd
from IPython.html.utils import path2url
from .common import INKSCAPE_PATH, FOLDER, NOTEBOOK_DIR, BUILD_ROOT_FOLDER

BUILD_FOLDER = ""
BUILD_FULL_PATH = ""


def convert_svg_to_png(svg_fn):
    abs_svg_fn = path.abspath(path.join(BUILD_FULL_PATH, svg_fn))
    abs_png_fn = abs_svg_fn.replace(".svg", ".png")
    print "converting ", abs_svg_fn, abs_png_fn

    subprocess.call([INKSCAPE_PATH, '-d', '300',
             '-f', abs_svg_fn,
             '-e', abs_png_fn], shell=True)


def to_md_table(df_in):

    def _f():
        df = df_in.astype(str)
        df.index = df.index.astype(str)
        name = df.index.name
        df = df.reset_index()
        n = list(df.apply(lambda s:s.str.len().max() + 4))


        split_line = "+" + "+".join(["-" * x for x in n]) + "+"
        yield split_line

        columns = list(df.columns)
        #columns[0] = ""
        yield "|" + "|".join(c.center(x) for x, c in zip(n, columns)) + "|"
        #if name is not None:
        #    yield split_line
        #    columns = [""] * len(n)
        #    columns[0] = name
        #    yield "|" + "|".join(c.center(x) for x, c in zip(n, columns)) + "|"

        yield split_line.replace("-", "=")

        for idx, row in df.iterrows():
            yield "|" + "|".join(c.center(x) for x, c in zip(n, row)) + "|"
            yield "+" + "+".join(["-" * x for x in n]) + "+"

    return "\n".join(_f())


def html_table_to_markdown(table):
    import io
    f = io.StringIO(table)
    df2 = pd.read_html(f, index_col=0)[0]
    s = df2.isnull().all()
    try:
        name = [x for x in s.index[s] if not x.startswith("Unnamed")][0]
        df2.index.name = name
    except:
        pass
    df2 = df2.dropna(1)
    return to_md_table(df2)


class GraphMerge(object):
    def __init__(self, rows, cols):
        self.rows, self.cols = rows, cols
        self.title = u""
        self.graphs = []

    def append(self, graph):
        self.graphs.append(graph)

    def is_completed(self):
        return len(self) == self.rows * self.cols

    def __len__(self):
        return len(self.graphs)

    def do_merge(self):
        import cv2
        import numpy as np

        images = [cv2.imread(path.join(BUILD_FULL_PATH, graph)) for graph in self.graphs]
        shape = self.rows, self.cols
        try:
            widths = np.array([img.shape[1] for img in images]).reshape(shape).max(axis=0)
            heights = np.array([img.shape[0] for img in images]).reshape(shape).max(axis=1)
        except AttributeError:
            import IPython
            IPython.embed()
        width = widths.sum()
        height = heights.sum()

        merge_img = np.zeros((height, width, 3), np.uint8)

        cumsum_width = np.r_[0, np.cumsum(widths)]
        cumsum_height = np.r_[0, np.cumsum(heights)]

        for i, img in enumerate(images):
            row = i // self.cols
            col = i % self.cols
            y = cumsum_height[row]
            x = cumsum_width[col]
            h, w = img.shape[:2]
            merge_img[y:y+h, x:x+w] = img

        cv2.imwrite(path.join(BUILD_FULL_PATH, self.graphs[-1].replace(".png", ".merge.png")), merge_img)


class Worker(object):

    def __init__(self):
        self.svg_graphs = []
        self.graph_merges = []
        self.ignore_next_figure = False
        self.hide_output = False

    def process_graph(self, output, ext):
        if self.ignore_next_figure:
            self.ignore_next_figure = False
            return u""

        if ext == "svg":
            graph_path = output.metadata.filenames["image/svg+xml"]
            self.svg_graphs.append(graph_path)
            graph_path = graph_path.replace(".svg", ".png")
        elif ext == "png":
            graph_path = output.metadata.filenames["image/png"]
        elif ext == "jpg":
            graph_path = output.metadata.filenames["image/jpeg"]

        title = output.get('title', ext)

        if self.graph_merges and not self.graph_merges[-1].is_completed():
            self.graph_merges[-1].append(graph_path)
            if len(self.graph_merges[-1]) == 1:
                self.graph_merges[-1].title = output.get('title', ext)
            if self.graph_merges[-1].is_completed():
                graph_path = graph_path.replace(".png", ".merge.png")
                title = self.graph_merges[-1].title

        md_code = u"![{title}]({filename})".format(
            title=title,
            filename=graph_path
        )

        if not self.graph_merges or self.graph_merges[-1].is_completed():
            return md_code
        else:
            return u""

    def process_code(self, code):
        lang = u"python"
        result = []
        if code.strip().startswith(u"#%figonly="):
            return u""

        if code.strip().startswith(u"%%dot"):
            return u""

        hide_flag = False
        lines = code.split(u"\n")
        if not lines:
            return ""

        first_line = lines[0]
        for line in lines:
            if line.startswith((u"#%fig", u"#%nofig", u"%%disabled", "%C ", "%%include", "%%language")):
                continue
            if line.startswith(u"#%hide_output"):
                self.hide_output = True
                continue
            if line.startswith(u"%omit"):
                line = line.replace(u"%omit ", u"")
            if line.startswith("%col "):
                line = re.match(ur"%col\s+\d+\s+(.+)", line).group(1)
            if line.startswith("#%hide"):
                hide_flag = True
            if line.startswith("#%show"):
                hide_flag = False
                continue
            if not hide_flag:
                result.append(line)

        code = u"\n".join(result)

        for num in u"❶❷❸❹❺❻❼❽❾":
            code = code.replace(u"#" + num, num)
            code = code.replace(u"# " + num, num)


        code = code.rstrip()

        if not first_line.startswith("%%language"):
            code = code.rstrip(";")

        if code.strip():
            return u"```{lang}\n##CODE\n{code}\n```".format(lang=lang, code=code)
        else:
            return u""

    def process_stream(self, output):
        if self.hide_output:
            return ""

        return u"```\n##OUTPUT\n{}\n```".format(output.text)

    def process_output(self, output):
        if self.hide_output:
            return ""

        try:
            text = output.text
        except AttributeError:
            text = output["data"]["text/plain"]

        text = text.strip()
        if text == "<IPython.core.display.Javascript object>":
            return ""
        if text:
            return "```\n##OUTPUT\n{}\n```".format(text)
        else:
            return ""

    def process_input(self, cell):
        outputs = [item for item in cell.outputs if "data" in item and "image/" in "|".join(item.data.keys())]
        idx = 0
        lines = cell.source.split(u"\n")
        self.hide_output = False
        if lines and lines[0].startswith(("%%file", "%%include")):
            self.hide_output = True
        for line in lines:
            if line.startswith(u"#%nofig"):
                self.ignore_next_figure = True
            if line.startswith(u"#%fig"):
                if idx < len(outputs):
                    outputs[idx][u"title"] = line[line.index("=")+1:].strip()
                    idx += 1
            if line.startswith(u"#%fig["):
                rows, cols = map(int, line[6:9].split(u"x"))
                merge = GraphMerge(rows, cols)
                self.graph_merges.append(merge)

        return self.process_code(cell.source)

    def process_html(self, html):
        if self.hide_output:
            return ""

        if "Generated by Cython" in html:
            return ""
        if "<table" in html:
            return html_table_to_markdown(html)
        return html

    def process_markdown(self, md):
        md_strip = md.strip()
        if md_strip.startswith(u"---") and md_strip.endswith(u"---"):
            lines = md_strip.split(u"\n")
            return u"%BLOCK_START\n{}\n%BLOCK_END".format(u"\n".join(lines[1:-1]))
        elif md_strip.startswith(u"> http"):
            return u"%LINK\n\n{}".format(md)
        else:
            res = re.match(r"> \*\*(SOURCE|TIP|WARNING|QUESTION|LINK)\*\*", md_strip)
            if res is not None:
                return u"%{}\n\n{}".format(res.group(1), u"\n".join(md_strip.split(u"\n")[2:]))
        return md

def read_notebook(notebook_name):
    from IPython.nbformat import read
    with open(notebook_name, "rb") as f:
        book = read(f, 4)

    #remove cells before heading cell
    for i, cell in enumerate(book.cells):
        if cell.cell_type == "markdown" and cell.source.startswith("#"):
            break

    del book.cells[:i]

    notebook_name_path = u"/".join(notebook_name.split(u"\\")[-2:])

    cell = {u'source': u'''> **SOURCE**

> 与本节内容对应的Notebook为：`{}`'''.format(notebook_name_path),
            u'cell_type': u'markdown', u'metadata': {}}

    if len(book.cells) >= 10:
        book.cells.insert(1, cell)

    return book


def convert_nb_to_markdown(notebook_path):
    notebook_path = path.abspath(notebook_path)
    notebook_name = path.basename(notebook_path)
    notebook_base_name = notebook_name.replace(".ipynb", "")
    exporter = MarkdownExporter(template_path=[FOLDER], template_file="mk.tpl")
    filter_data_type = exporter.environment.filters['filter_data_type']
    filter_data_type.display_data_priority.append("text/markdown")

    writer = FilesWriter(build_directory=BUILD_FULL_PATH)
    in_resources = {
        'unique_key':notebook_base_name,
        'output_files_dir': "%s_files" % notebook_base_name,
        'worker': Worker()
    }

    book = read_notebook(notebook_path)

    output, resources = exporter.from_notebook_node(book, in_resources)
    writer.write(output, resources, notebook_name=notebook_base_name)

    worker = resources["worker"]

    for svg_fn in worker.svg_graphs:
        convert_svg_to_png(svg_fn)

    for merge in worker.graph_merges:
        merge.do_merge()

    def replace_images(g):
        fn = g.group(1)
        src = path.join(NOTEBOOK_DIR, "images", fn)

        dst_folder = path.join(BUILD_FULL_PATH, in_resources['output_files_dir'])
        if not path.exists(dst_folder):
            os.mkdir(dst_folder)

        dst = path.join(dst_folder, fn)

        shutil.copyfile(src, dst)
        return "{}/{}".format(path.basename(path.dirname(dst)), path.basename(dst))

    def replace_svg(g):
        svg_fn = g.group(1)
        build_folder = BUILD_FULL_PATH
        abs_svg_fn = path.abspath(path.join(build_folder, svg_fn))
        abs_png_fn = abs_svg_fn.replace(".svg", ".png")
        print "converting ", abs_svg_fn, abs_png_fn

        subprocess.call([INKSCAPE_PATH, '-d', '300',
                 '-f', abs_svg_fn,
                 '-e', abs_png_fn], shell=True)

        return "({})".format(svg_fn.replace(".svg", ".png"))

    md_fn = path.join(BUILD_FULL_PATH, notebook_base_name + ".md")

    with open(md_fn, "rb") as f:
        text = f.read()
        text = re.sub(r"/files/images/(.+\.png)", replace_images, text)
        image_folder = notebook_base_name + "_files"
        #text = re.sub(r"\((.+\.svg)\)", replace_svg, text)

    with open(md_fn, "wb") as f:
        f.write(text)

    return md_fn


def iter_notebooks(folder, pattern="*.ipynb"):
    for fn in glob(path.join(folder, pattern)):
        full_path = path.abspath(fn)
        base_name = path.basename(full_path)
        if base_name.startswith("_"):
            continue
        if re.match(r"\w+-\w\d\d-.+?\.ipynb", base_name):
            yield full_path


def convert_folder_to_md(folder, pattern="*.ipynb"):
    global BUILD_FOLDER, BUILD_FULL_PATH

    BUILD_FOLDER = "build_" + folder
    BUILD_FULL_PATH = path.join(BUILD_ROOT_FOLDER, BUILD_FOLDER)

    build_folder = BUILD_FULL_PATH
    print "build_folder", build_folder

    if path.exists(build_folder):
        shutil.rmtree(build_folder)
    os.mkdir(build_folder)

    md_files = []
    for notebook in iter_notebooks(path.join(NOTEBOOK_DIR, folder), pattern):
        print notebook
        md_files.append(convert_nb_to_markdown(notebook))

    md_text = []
    for md_file in md_files:
        with open(md_file, "rb") as f:
            md_text.append(f.read())

    final_md_file = path.join(build_folder, folder + ".md")

    with open(final_md_file, "wb") as f:
        f.write("\n\n".join(md_text))

    return  final_md_file


def convert_md_to_docx(md_filename):
    current_dir = os.getcwdu()
    os.chdir(path.dirname(md_filename))
    cmd = "pandoc --no-highlight -s -S --from=markdown --to=docx {infile} -o {outfile}".format(
        infile = md_filename,
        outfile = md_filename.replace(".md", ".docx")
    )
    subprocess.call(cmd)
    os.chdir(current_dir)

def test_merge_graphs():
    global BUILD_FOLDER
    BUILD_FOLDER = u"build_04-matplotlib"
    merge = GraphMerge(1, 3)
    merge.graphs = [u'matplotlib-100-fastdraw_files/matplotlib-100-fastdraw_27_0.png', u'matplotlib-100-fastdraw_files/matplotlib-100-fastdraw_29_0.png', u'matplotlib-100-fastdraw_files/matplotlib-100-fastdraw_31_0.png']
    merge.do_merge()

if __name__ == '__main__':
    #md_file = convert_folder_to_md("04-matplotlib", "matplotlib-100-fastdraw.ipynb")
    #md_file = convert_folder_to_md("02-numpy", pattern="numpy-*.ipynb")
    md_file = path.abspath(r"build_02-numpy\02-numpy.md")
    convert_md_to_docx(md_file)
