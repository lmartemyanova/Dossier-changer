import os

from dossier_administration.form_dossier import form_dossier
# from dossier_administration.form_xml import form_xml


def show_info():
    print('''
    Список доступных команд:
    "fd" - сформировать досье (вам будет последовательно предложено ввести данные для автоматического сбора шаблонов)
    "xml" - сформировать XML-файл
    ''')
    return


def manage_program():
    while True:
        command = input("Введите команду: ")
        if command == "fd":
            folder = input("Введите путь к папке, в которой находятся шаблоны для изменения (или поместите папку в "
                           "корневую папку проекта): ")
            folder = os.path.abspath(folder)
            if folder == '':
                folder = os.getcwd()
            form_dossier(folder=folder)
        # elif command == 'xml': folder = input("Введите путь к папке с документами для формирования XML (или
        # поместите папку в корневую папку проекта): ") if folder == '': folder = os.getcwd() form_xml(folder)
