import os.path
import win32com.client
baseDir = os.getcwd()

word = win32com.client.Dispatch("Word.application")
word.visible = False

for dir_path, dirs, files in os.walk(baseDir):
	for file_name in files:
		file_path = os.path.join(dir_path, file_name)
		file_name, file_extension = os.path.splitext(file_path)
		if file_extension.lower() == '.doc': #
			docx_file = '{0}{1}'.format(file_path, 'x')
			if not os.path.isfile(docx_file): # Skip conversion where docx file already exists
				print('Converting: {0}'.format(file_path))
				try:
					wordDoc = word.Documents.Open(file_path, True, True, True)
					wordDoc.SaveAs2(docx_file[:-5], 16)
					wordDoc.Close()
				except Exception as e:
					print('Failed to Convert: {0}'.format(file_path))
					print(e)






# pip install -U pypiwin32
from win32com import client as wc

import os
w = wc.Dispatch('Word.Application')

paths = []
print(os.getcwd())
folder = os.getcwd()
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('doc') and not file.startswith('~'):
            paths.append(os.path.join(root, file))

for path in paths:
    try:
        doc = w.Documents.Open(path, False, False, False)
        doc.SaveAs('{0}{1}'.format(path, 'x'), 16)
        doc.Close()
    except Exception as e:
        print('Failed to Convert: {0}'.format(path))
        print(e)
w.Quit()


from win32com import client as wc
import os
import re
import sys
import comtypes.client

paths = []
folder = os.path.join(os.getcwd(), 'dossier')
for root, dirs, files in os.walk(folder):
    for file in files:
        filename, file_extension = os.path.splitext(file)
        # if file.endswith('doc') and not file.startswith('~'):
        #     paths.append(os.path.join(root, file))
        if file_extension.lower() == ".doc":
            paths.append(os.path.join(root, file))
            # save_as_docx(file_conv)
            # print("%s ==> %sx" %(file_conv,f))

for path in paths:
    w = wc.gencache.EnsureDispatch('Word.Application')
    doc = w.Documents.Open(path)
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    doc.SaveAs(new_file_abs, FileFormat=16)
    doc.Close()
w.Quit()
# _________________
# for path in paths:
#     w = wc.gencache.EnsureDispatch('Word.Application')
#     # w = win32.gencache.EnsureDispatch('Word.Application')
#     doc = w.Documents.Open(path)
#     new_file_abs = os.path.abspath(path)
#     new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
#     doc.SaveAs(new_file_abs, 16)
#     doc.Close(wdSaveChanges)
#     w.Quit()
# import re
# import os
# import sys
# import win32com.client as win32
# from win32com.client import constants
#
# path = os.path.join(os.getcwd(), 'dossier')
# def save_as_docx(path):
#     # Opening MS Word
#     word = win32.gencache.EnsureDispatch('Word.Application')
#     doc = word.Documents.Open(path)
#     doc.Activate ()
#
#     # Rename path with .docx
#     new_file_abs = os.path.abspath(path)
#     new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
#
#     # Save and Close
#     word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
#     doc.Close(False)
#
# def main():
#     source = ABS_PATH
#
#     for root, dirs, filenames in os.walk(source):
#         for f in filenames:
#             filename, file_extension = os.path.splitext(f)
#
#             if file_extension.lower() == ".doc":
#                 file_conv = os.path.join(root, f)
#                 save_as_docx(file_conv)
#                 print("%s ==> %sx" %(file_conv,f))
#
# if __name__ == "__main__":
#     main()

# __________________________________
# from glob import glob
# import re
# import os
# import win32com.client as win32
# from win32com.client import constants
#
#
# def save_as_docx(path):
#
#     # Открываем Word
#     try:
#         word = win32.gencache.EnsureDispatch('Word.Application')
#         doc = word.Documents.Open(path)
#         doc.Activate()
#
#         # Меняем расширение на .docx и добавляем в путь папку
#         # для складывания конвертированных файлов
#         new_file_abs = str(os.path.abspath(path)).split("\\")
#         new_dir_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}"
#         new_file_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}\\doc_convert\\{new_file_abs[2]}"
#         new_file_abs = os.path.abspath(new_file_abs)
#         if not os.path.isdir(f'{new_dir_abs}\\doc_convert'):
#             os.mkdir(f'{new_dir_abs}\\doc_convert')
#         new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
#         print(new_file_abs)
#
#         # Сохраняем и закрываем
#         word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
#         doc.Close(False)
#     except:
#         return str(path).split("\\")[-1]
#
#
# def path_doc(paths):
#     dict_error_file = []
#     for path in paths:
#         err = save_as_docx(path)
#         if err != None:
#             dict_error_file.append(err)
#     if len(dict_error_file) >= 1:
#         print(f'\nНе конвертированные файлы (ошибка открытия - файл поврежден):\n{dict_error_file}')
#
#
# def main():
#     dirs = input(f'Введите путь к папке с файлами\n(пример: C:\\temp) >>> ')
#     paths = glob(f'{dirs}\\*.doc', recursive=True)
#     path_doc(paths)
#
#
# if __name__ == "__main__":
#     main()
