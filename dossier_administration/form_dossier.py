from filetypes.word_docx import WordDocx, WordDoc
import os


def get_text_dict() -> dict:
    """
    To form the dict with the text, which need to be replaced in all files of the dossier if found, from user input.
    :return: dict of the text to be replaced
    """

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


def get_filenames_attributes() -> dict:
    """
    To form the dict with the attributes of the filenames, which need to be replaced, from user input.
    :return: dict of filenames attributes to be replaced
    """

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


def form_dossier(folder: str) -> None:
    """
    To form the dossier from files of the other product, by the replacement of the text.
    After the usage of this method the dossier has to be corrected and checked by hands before forming .xml!
    :param folder: the path to the folder with files to form dossier.
                   Only .doc/.docx and .pdf files are allowed, the others will ignore.
    :return: None
    """

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
