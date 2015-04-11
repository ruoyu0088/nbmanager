import sys
from collections import OrderedDict
import win32com.client
from win32com.client import constants as C
from os import path
import time
from mdcreater import NOTEBOOK_DIR
from diff_editor import diff_html

name = sys.argv[1]

folder = path.join(NOTEBOOK_DIR, name)
docx_file = path.join(folder, name + ".docx")

app = win32com.client.gencache.EnsureDispatch("Word.Application")
app.Visible = True

doc = app.Documents.Open(docx_file)

app.ActiveWindow.View.ShowRevisionsAndComments = True
app.ActiveWindow.View.RevisionsView = C.wdRevisionsViewFinal
time.sleep(1.0)

paragraphs = []

for idx, rev in enumerate(doc.Revisions):
    r = rev.Range
    r.Select()
    p = app.Selection.Range.Paragraphs(1)
    paragraphs.append(p)

app.ActiveWindow.View.ShowRevisionsAndComments = False
time.sleep(1.0)

final_text = []
for p in paragraphs:
    final_text.append(p.Range())


app.ActiveWindow.View.RevisionsView = C.wdRevisionsViewOriginal
time.sleep(1.0)

original_text = []
for p in paragraphs:
    original_text.append(p.Range())

changes = OrderedDict((otext, ftext) for (otext, ftext) in
                        zip(original_text, final_text) if otext != ftext)

import pickle
with open(path.join(folder, "changes.pickle"), "wb") as f:
    pickle.dump(changes, f)

with open(path.join(folder, "changes.html"), "w") as f:
    f.write("<html><body>")
    for otext, ftext in changes.iteritems():
        f.write("<p>")
        f.write(diff_html(otext, ftext).encode("utf8"))
        f.write("</p>")
    f.write("</body></html>")

doc.Close()
app.Quit()