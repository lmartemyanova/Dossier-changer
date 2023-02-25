import docx
import os

Dictionary = {"Лираглутид, таблетки, 0,1 мг": "Эзомепразол, капсулы кишечнорастворимые, 20 мг, 50 мг",
              'ООО «Компания»': 'ООО «Другая компания»',
              "1.2.1. Какой-то раздел досье": "Какой-то другой раздел досье"}

paths = []
folder = os.getcwd()
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('docx') and not file.startswith('~'):
            paths.append(os.path.join(root, file))

for path in paths:
    doc = docx.Document(path)
    style = doc.styles['Normal']
    font = style.font

    for i in Dictionary:
        for paragraph in doc.paragraphs:
            if paragraph.text.find(i) >= 0:
                paragraph.text = paragraph.text.replace(i, Dictionary[i])

    for j in Dictionary:
        for table in doc.tables:
            for col in table.columns:
                for cell in col.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.find(j) >= 0:
                            paragraph.text = paragraph.text.replace(j, Dictionary[j])

    for k in Dictionary:
        for section in doc.sections:
            header = section.header

            for paragraph in header.paragraphs:
                if paragraph.text.find(k) >= 0:
                    paragraph.text = paragraph.text.replace(k, Dictionary[k])

            for table in header.tables:
                for col in table.columns:
                    for cell in col.cells:
                        for paragraph in cell.paragraphs:
                            if paragraph.text.find(k) >= 0:
                                paragraph.text = paragraph.text.replace(k, Dictionary[k])

    doc.save(os.path.basename(path))