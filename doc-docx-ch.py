from win32com import client as wc

import os
w = wc.Dispatch('Word.Application')

paths = []
folder = os.getcwd()
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('doc') and not file.startswith('~'):
            paths.append(os.path.join(root, file))

for path in paths:
    doc = w.Documents.Open(path)
    doc.SaveAs(path+"x", 16)
    doc.Close()

w.Quit()