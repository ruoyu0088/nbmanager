#encoding=utf8
import difflib
import json
import pickle
from glob import glob
from os import path
from PyQt4 import QtGui
from IPython.nbformat import read
from .mdcreater import NOTEBOOK_DIR, iter_notebooks
#font_db = QtGui.QFontDatabase()
#print u"\n".join([unicode(s) for s in font_db.families()])


DEL_FORMAT =  u'<span style="text-decoration: line-through; color:red; font-weight:bold;">{}</span>'
INSERT_FORMAT = u'<u style="color:orange;  font-weight:bold;">{}</u>'


def strip_ascii(text):
    return u"".join(c for c in text if 0x4e00 <= ord(c) <= 0x9FCC and c not in u"图")

def strip_invisible(text):
    return u"".join(c for c in text if ord(c) >= 32)

def diff_html(otext, ftext):
    html1 = ""
    sm = difflib.SequenceMatcher(None, otext, ftext)
    for code, i1, i2, j1, j2 in sm.get_opcodes():
        if code == "equal":
            html1 += otext[i1:i2]
        elif code == "delete":
            html1 += DEL_FORMAT.format(otext[i1:i2])
        elif code == "replace":
            html1 += DEL_FORMAT.format(otext[i1:i2])
            html1 += INSERT_FORMAT.format(ftext[j1:j2])
        elif code == "insert":
            html1 += INSERT_FORMAT.format(ftext[j1:j2])
    return html1.replace("\n", "<br>")


def iter_notebook_paragraphs(folder):
    number, name = path.basename(folder).split("-")
    for notebook in iter_notebooks(folder):
        basename = path.basename(notebook)
        with open(notebook, "rb") as f:
            book = read(f, 4)
        for i, cell in enumerate(book.cells):
            text = cell.source
            yield basename, i, text


def auto_modify_text(otext, ftext, text):
    lines = text.split("\n")
    best_line = max(lines, key=lambda line:difflib.SequenceMatcher(None, otext, line).ratio())
    if abs(len(otext) - len(best_line)) > 40:
        return text

    index = lines.index(best_line)

    a, b = best_line, ftext
    sm = difflib.SequenceMatcher(None, a, b)

    res = ""

    keep = "`$*"

    for code, i1, i2, j1, j2 in sm.get_opcodes():
        if code == "equal":
            res += a[i1:i2]
        elif code == "delete":
            to_del = a[i1:i2]
            if "ref:" in to_del:
                res += to_del
            else:
                res += u"".join(c for c in to_del if c in keep)
        elif code == "replace":
            to_del = a[i1:i2]
            if "ref:" in to_del:
                res += to_del
            else:
                res += u"".join(c for c in to_del if c in keep)
                res += b[j1:j2]
        elif code == "insert":
            res += b[j1:j2]

    lines[index] = res
    return "\n".join(lines)


class DiffEditor(QtGui.QWidget):

    def __init__(self):
        super(DiffEditor, self).__init__()
        hbox = QtGui.QHBoxLayout()
        self.folder_combobox = QtGui.QComboBox(self)
        self.load_button = QtGui.QPushButton("Load", self)
        self.prev_button = QtGui.QPushButton("&Prev", self)
        self.next_button = QtGui.QPushButton("&Next", self)
        self.original_button = QtGui.QPushButton("&Original", self)
        self.save_button = QtGui.QPushButton("&Save", self)
        self.info = QtGui.QLabel("", self)

        hbox.addWidget(self.folder_combobox)
        hbox.addWidget(self.load_button)
        hbox.addWidget(self.prev_button)
        hbox.addWidget(self.next_button)
        hbox.addWidget(self.original_button)
        hbox.addWidget(self.save_button)
        hbox.addWidget(self.info)

        self.load_button.clicked.connect(self.on_load)
        self.prev_button.clicked.connect(self.on_prev)
        self.next_button.clicked.connect(self.on_next)
        self.original_button.clicked.connect(self.on_original)
        self.save_button.clicked.connect(self.on_save)
        box = QtGui.QVBoxLayout()

        self.html_viewer = QtGui.QTextEdit(self)
        self.html_viewer.setReadOnly(True)
        self.html_viewer.setMaximumHeight(150)
        self.html_viewer2 = QtGui.QTextEdit(self)
        self.html_viewer2.setReadOnly(True)

        self.cell_editor = QtGui.QTextEdit(self)
        self.cell_editor.textChanged.connect(self.on_text_changed)

        font = QtGui.QFont()
        font.setPointSize(14)
        font.setFamily(u"Yahei Mono")
        self.html_viewer.setFont(font)
        self.html_viewer2.setFont(font)
        self.cell_editor.setFont(font)
        self.info.setFont(font)

        box.addLayout(hbox)
        box.addWidget(self.html_viewer)
        box.addWidget(self.html_viewer2)
        box.addWidget(self.cell_editor)
        self.setLayout(box)

        self.populate_folder()

    def populate_folder(self):
        for folder in glob(path.join(NOTEBOOK_DIR, "??-*")):
            if path.exists(path.join(folder, "changes.pickle")):
                self.folder_combobox.addItem(path.basename(folder))

    def closeEvent(self, evt):
        self.on_save()

    def on_load(self):
        self.load_folder(path.join(NOTEBOOK_DIR, str(self.folder_combobox.currentText())))

    def on_save(self):
        self.db[self.db_key] = self.editor_text
        self.cell_editor.setStyleSheet("background-color:#ffffff;")
        with open(self.db_name, "wb") as f:
            json.dump(self.db, f)

    def on_original(self):
        res = QtGui.QMessageBox.question(self, "Set to Original", "Are you sure?",
                                         QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        if res == QtGui.QMessageBox.Yes:
            self.cell_editor.setPlainText(self.cell_text)

    def on_prev(self):
        self.index = max(0, self.index - 1)
        self.show_change()

    def on_next(self):
        self.index = min(len(self.changes) - 1, self.index + 1)
        self.show_change()

    def on_text_changed(self):
        self.show_editor_change()
        if self.db_key is not None:
            self.db[self.db_key] = self.editor_text

    @property
    def db_name(self):
        return path.join(self.folder, "changes.json")

    @property
    def db_key(self):
        text = str(self.info.text())
        try:
            return text.split(":")[1]
        except IndexError:
            return None

    @property
    def editor_text(self):
        return unicode(self.cell_editor.toPlainText())

    def load_folder(self, folder):
        self.load_notebooks(folder)
        self.load_changes(path.join(folder, "changes.pickle"))

    def load_notebooks(self, folder):
        self.folder = folder
        self.cells = list(iter_notebook_paragraphs(folder))
        if path.exists(self.db_name):
            with open(self.db_name, "rb") as f:
                self.db = json.load(f)
        else:
            self.db = {}

    def load_changes(self, fn):
        with open(fn, "rb") as f:
            self.changes = [(strip_invisible(k), strip_invisible(v)) for k, v in pickle.load(f).items()]
            self.index = 0
            self.show_change()

    def show_editor_change(self):
        html = diff_html(self.cell_text, self.editor_text)
        self.html_viewer2.clear()
        self.html_viewer2.insertHtml(html)

    def show_change(self):
        otext, ftext = self.changes[self.index]
        otext_strip = strip_ascii(otext)
        html = diff_html(otext, ftext)
        self.html_viewer.clear()
        self.html_viewer.insertHtml(html)

        for name, index, cell_text in self.cells:
            cell_text_strip = strip_ascii(cell_text)
            if otext_strip in cell_text_strip:
                break
        else:
            name = "-"
            index = -1
            cell_text = "Can't find the cell"

        self.info.setText("{}:{} - {}".format(self.index, name, index))

        if index >= 0:
            self.cell_text = cell_text
            if self.db_key in self.db:
                cell_text_modified = self.db[self.db_key]
                self.cell_editor.setStyleSheet("background-color:#ffffff;")
            else:
                cell_text_modified = auto_modify_text(otext, ftext, cell_text)
                self.cell_editor.setStyleSheet("background-color:#cccccc;")
                #self.db[self.db_key] = cell_text_modified
            self.cell_editor.clear()
            self.cell_editor.setPlainText(cell_text_modified)
            self.show_editor_change()
        else:
            self.cell_editor.setPlainText(cell_text)




if __name__ == '__main__':
    import sys
    app = QtGui.QApplication(sys.argv)
    editor = DiffEditor()
    #editor.load_notebooks(r"C:\Users\AAAF929\Documents\python\notebooks\docxmanager\02-numpy")
    #editor.load_changes(r"C:\Users\AAAF929\Documents\python\notebooks\docxmanager\changes.pickle")
    editor.show()
    app.exec_()
