from os import path
import subprocess
import threading
import time
from traits.api import HasTraits, List, Str, Enum, Bool, Button, Code, Int
from traitsui.api import View, Item, EnumEditor, HGroup, VGroup, CodeEditor
from pyface.timer.api import Timer
from .common import iter_chapters, NOTEBOOK_DIR, iter_notebooks


class DocxBuildGUI(HasTraits):

    chapters = List(Str)
    selected_chapter = Str

    notebooks = List(Str)
    selected_notebook = Str
    all_notebooks = Bool(True)

    build_step = Enum("all", ["md", "docx", "pdocx", "all"])
    build_button = Button("Build")

    stdout = Code()
    selected_line = Int

    view = View(
        VGroup(
            HGroup(
                Item("selected_chapter", editor=EnumEditor(name="object.chapters")),
                Item("all_notebooks"),
                Item("selected_notebook", editor=EnumEditor(name="object.notebooks"),
                     enabled_when="not object.all_notebooks"),
                Item("build_step"),
                Item("build_button", show_label=False),
            ),
            Item("stdout", show_label=False,
                 editor=CodeEditor(lexer="text", selected_line="selected_line"))
        ),
        title="Docx builder",
        width=800,
        height=600,
        resizable=True,
    )

    def __init__(self):
        self.process = None
        self.timer = Timer(500, self.select_last_line)
        self.chapters = [path.basename(p) for p in iter_chapters()]

    def _selected_chapter_changed(self):
        folder = path.join(NOTEBOOK_DIR, self.selected_chapter)
        self.notebooks = [path.basename(nb) for nb in iter_notebooks([folder])]

    def _build_button_fired(self):
        args = ["python", "-m", "nbmanager", self.selected_chapter]
        if not self.all_notebooks:
            args.extend(["-p", self.selected_notebook])
        args.extend(["-s", self.build_step])
        self.stdout = ""
        self.process = subprocess.Popen(args, stdout=subprocess.PIPE)
        self.stdout_thread = threading.Thread(target=self.check_pipe)
        self.stdout_thread.start()

    def check_pipe(self):
        while self.process.poll() is None:
            self.stdout += self.process.stdout.readline()
            time.sleep(0.1)
        self.stdout += self.process.stdout.read()

    def select_last_line(self):
        self.selected_line = self.stdout.count("\n") + 1


if __name__ == '__main__':
    gui = DocxBuildGUI()
    gui.configure_traits()