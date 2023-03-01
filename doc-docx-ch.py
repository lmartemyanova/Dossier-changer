import aspose.words as aw
import os

paths = []
folder = os.getcwd()

for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('doc') and not file.startswith('~'):
            paths.append(os.path.join(root, file))

for path in paths:
    doc = aw.Document(path)
    doc.save(path+'x')