from word_docx import WordDocx, WordDoc
import os


def show_info():
    print('''
    Список доступных команд:
    "fd" - сформировать досье (вам будет последовательно предложено ввести данные для автоматического сбора шаблонов)
    "xml" - сформировать XML-файл
    ''')
    return


def get_text_dict():
    text_dict = {}
    print('''Ниже вам будет предложена форма для ввода текста, который требуется заменить в каждом файле досье, /
    таких как название препарата, лекарственная форма, производители ГЛФ и АФС и т.д.
    Введите ниже текст, который вы хотите заменить в формате "старый текст" - enter - "новый текст".
    Во избежание ошибок вводите сразу фразы целиком (например, "Наименование ЛП, лекарственная форма, дозировка").
    Вы можете ввести несколько фраз для замены последовательно (попарно старый текст-новый текст).
    Чтобы продолжить ввод фраз для замены, введите "Y".
    Чтобы закончить ввод текста и сформировать окончательный список замен, введите "N".
    ''')
    while True:
        one_more_pair = input('Вы хотите ввести текст для замены? Введите Y или N: ')
        if one_more_pair == 'Y' or one_more_pair == 'y':
            old = input('Введите текст, который требуется заменить во всех документах: ')
            new = input('Введите замещающий текст: ')
            text_dict[old] = new
        elif one_more_pair == 'N' or one_more_pair == 'n':
            break
        else:
            print('Кажется, вы ввели что-то не так. Попытайтесь снова. ')
    return text_dict


def get_filenames_attributes():
    filenames_attributes = {}
    print('''Введите атрибуты АФС или ГЛФ, которые необходимо заменить в названии(!) файлов (документов).
    Обратите внимание на необходимость полного соответствия исходных атрибутов (в т.ч. символы ".", "_", "-". 
    Если необходимо ввести еще одну пару атрибутов, напишите "Y".
    Если вы хотите закончить ввод атрибутов и сформировать окончательный список замен, введите "N"''')
    while True:
        one_more_pair = input('Вы хотите ввести атрибуты для замены? Введите Y или N: ')
        if one_more_pair == 'Y' or one_more_pair == 'y':
            old = input('Введите текст, который требуется заменить во всех документах: ')
            new = input('Введите замещающий текст: ')
            filenames_attributes[old] = new
        elif one_more_pair == 'N' or one_more_pair == 'n':
            break
        else:
            print('Кажется, вы ввели что-то не так. Попытайтесь снова. ')
    return filenames_attributes


def form_dossier(folder):
    text_dict = get_text_dict()
    filenames_attributes = get_filenames_attributes()
    for root, dirs, files in os.walk(folder):
        for file in files:
            path = os.path.join(folder, file)
            if file.endswith('doc'):
                filename = WordDoc(file)
                filename.save_docx(path)
                filename.replace_text(text_dict, path)
                if '3.2.S.' or '2.3.S.' in filename:
                    filename.rename(filenames_attributes, path)
            elif file.endswith('docx'):
                filename = WordDocx(file)
                filename.replace_text(text_dict, path)
                if '3.2.S.' or '2.3.S.' in filename:
                    filename.rename(filenames_attributes, path)
            # elif file.endswith('pdf'):
            #     filename = PdfTxt(file)
    return


# def form_xml(folder):


if __name__ == '__main__':
    show_info()
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
