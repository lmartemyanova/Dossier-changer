import os
from win32com import client as wc
import docx
import contextlib

class WordDocx:
    def __init__(self, file):
        self.file = file

    @contextlib.contextmanager
    def replace_words(self):
        try:
            yield {}
        except Exception:
            print(self.docname)
        finally:


