import os
from filetypes.word_docx import WordDocx, WordDoc


def find(folder:str, phrase:str) -> list:
    files_with_phrase = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            path = os.path.join(folder, file)
            if file.endswith('doc'):
                filename = WordDoc(file)
                filename.save_docx(path)
                n = filename.find_usages(path=folder, phrase=phrase)
                if n:
                    files_with_phrase.append(n)
            elif file.endswith('docx'):
                filename = WordDocx(file)
                n = filename.find_usages(path=folder, phrase=phrase)
                if n:
                    files_with_phrase.append(n)
    return files_with_phrase
