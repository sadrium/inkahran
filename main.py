import os
from PyQt5.QtWidgets import *
from PyQt5 import QtGui
import openpyxl

app = QApplication([])
win = QWidget()
win.resize(1000, 600)
win.setWindowTitle('Приложение для проветки счетов ИНКАХРАН')

btn_dir_cft = QPushButton("Папка с файлами ЦФТ")
btn_dir_cft.setFont(QtGui.QFont("Times", 12, QtGui.QFont.Bold))
btn_dir_cft.adjustSize()
lw_files_cft = QListWidget()

btn_dir_incas = QPushButton("Папка с файлами ИНКАХРАН")
btn_dir_incas.setFont(QtGui.QFont("Times", 12, QtGui.QFont.Bold))
btn_dir_incas.adjustSize()
lw_files_incas = QListWidget()

btn_instruction = QPushButton("Инструкция")
btn_instruction.setFont(QtGui.QFont("Times", 12, QtGui.QFont.Bold))
btn_instruction.adjustSize()
btn_check = QPushButton("Сверить")
btn_check.setFont(QtGui.QFont("Times", 12, QtGui.QFont.Bold))
btn_check.adjustSize()

main_col = QVBoxLayout()
row1 = QHBoxLayout()
row2 = QHBoxLayout()

row1.addWidget(btn_dir_cft)
row2.addWidget(lw_files_cft)

row1.addWidget(btn_dir_incas)
row2.addWidget(lw_files_incas)

main_col.addWidget(btn_instruction)
main_col.addLayout(row1, 20)
main_col.addLayout(row2, 20)
main_col.addWidget(btn_check)

win.setLayout(main_col)

win.show()


def popup():
    instruction = QMessageBox()
    instruction.setWindowTitle('Инструкция по использованию приложения')
    instruction.setText('''Для работы с программой у Вас на компьютере должны быть следующие файлы: 
    1) пины-тарифы.txt
    2) проводка из цфт за месяц (название любое). Проводку также сохранять с кодировкой UTF-8.
    3) файл текстовый, который прислал ИНКАХРАН (название любое)
    Для того, чтобы сделать последний файл нужно зайти в файл Excel, нажать "Сохранить как" и сохранить в формате TXT.
    Затем зайти в блокнот, нажать "Файл -> Сохранить как..." и в окне "Кодировка" выбрать UTF-8. (Это очень важно).
    Далее, воспользовавшись кнопками выбора папок, необходимо выбрать папки, 
    где у Вас лежат файлы из ЦФТ и созданные текстовые файл сверок.
    Из получившихся списков двойным щелчком выбрать нужные файлы и нажать кнопку "Сверить".
    Если Вы сделали все по инструкции, у Вас должен получиться файл "ПИНы - комиссии итог.txt" в папке, 
    где лежат проводки из ЦФТ за месяц.''')
    instruction.setFont(QtGui.QFont("Times", 12))
    instruction.setIcon(QMessageBox.Information)
    instruction.setStandardButtons(QMessageBox.Ok)
    instruction.resize(100, 100)
    instruction.exec_()


workdir_cft = ''


def filter_cft(files, extensions):
    result_cft = []
    for filename in files:
        for ext in extensions:
            if filename.endswith(ext):
                result_cft.append(filename)
    return result_cft


def chooseWorkdir_cft():
    global workdir_cft
    workdir_cft = QFileDialog.getExistingDirectory()


def showFilenamesList_cft():
    extensions = ['.txt']
    chooseWorkdir_cft()
    filenames = filter_cft(os.listdir(workdir_cft), extensions)
    lw_files_cft.clear()
    for filename in filenames:
        lw_files_cft.addItem(filename)


workdir_incas = ''


def filter_incas(files, extensions):
    result_incas = []
    for filename in files:
        for ext in extensions:
            if filename.endswith(ext):
                result_incas.append(filename)
    return result_incas


def chooseWorkdir_incas():
    global workdir_incas
    workdir_incas = QFileDialog.getExistingDirectory()


def showFilenamesList_incas():
    extensions = ['.xlsx']
    chooseWorkdir_incas()
    filenames = filter_incas(os.listdir(workdir_incas), extensions)
    lw_files_incas.clear()
    for filename in filenames:
        lw_files_incas.addItem(filename)


name_cft = ''


def event_cft(item_cft):
    global name_cft
    name_cft = item_cft.text()
    return name_cft


name_incas = ''


def event_incas(item_incas):
    global name_incas
    name_incas = item_incas.text()
    return name_incas


def making_list_cft():
    with open(workdir_cft.replace('/', '\\') + '\\' + name_cft, 'r', encoding='utf-8') as file:
        table = file.read()

    dict_table = {}

    table = table.split('\n')
    split_table = []
    col = []

    for line in table:
        line = line.split('\t')
        split_table.append(line)

    for i in range(len(split_table[0])):
        for j in range(1, len(split_table)):
            col.append(split_table[j][i])
        dict_table[split_table[0][i]] = col
        col = []

    dict_table['ПИН'] = dict_table['Назначение платежа']

    items_to_delete = ['Счет Дебет', 'Счет Кредит', 'Состояние', 'Получатель', 'Компания', 'Документ',
                       'Назначение платежа']
    for item in items_to_delete:
        dict_table.pop(item)

    PIN = []

    for each in dict_table['ПИН']:
        each = each.split(' ЗАЯВЛЕНО')
        PIN.append(each[0])

    dict_table['ПИН'] = PIN

    summa = []

    for each in dict_table['Сумма']:
        each = each.replace(',', '.')
        summa.append(each)

    dict_table['Сумма'] = summa

    with open(workdir_cft.replace('/', '\\') + '\\' + 'пин-тариф.txt', 'r', encoding='utf-8') as file:
        rates = file.read()

    rates = rates.split('\n')
    PIN = []

    for line in rates:
        line = line.split('\t')
        PIN.append(line)

    rates = {}
    PIN = PIN[1::]
    row = []
    rates['ПИН'] = ['Инкассация %', 'Инкассация мин', 'Пересчет %', 'Пересчет мин', 'НДС']

    for i in range(len(PIN)):
        for j in range(1, len(PIN[0])):
            row.append(PIN[i][j])
        rates[PIN[i][0]] = row
        row = []

    commission = {}

    for i in range(len(dict_table['ПИН'])):
        if dict_table['ПИН'][i] in rates.keys():
            if rates[dict_table['ПИН'][i]][0] == '0.00':
                commission[dict_table['ПИН'][i]] = float(rates[dict_table['ПИН'][i]][1]) * 1.2
            else:
                if float(dict_table['Сумма'][i]) * float(rates[dict_table['ПИН'][i]][0]) * 1.2 < float(
                        rates[dict_table['ПИН'][i]][1]) * 1.2:
                    commission[dict_table['ПИН'][i]] = float(rates[dict_table['ПИН'][i]][1]) * 1.2
                else:
                    commission[dict_table['ПИН'][i]] = float(dict_table['Сумма'][i]) * float(
                        rates[dict_table['ПИН'][i]][0]) * 1.2
            if rates[dict_table['ПИН'][i]][2] == '0.00':
                commission[dict_table['ПИН'][i]] += float(rates[dict_table['ПИН'][i]][3])
            else:
                if float(dict_table['Сумма'][i]) * float(rates[dict_table['ПИН'][i]][2]) < float(
                        rates[dict_table['ПИН'][i]][3]):
                    commission[dict_table['ПИН'][i]] += float(rates[dict_table['ПИН'][i]][3])
                else:
                    commission[dict_table['ПИН'][i]] += float(dict_table['Сумма'][i]) * float(
                        rates[dict_table['ПИН'][i]][2])

    for key in commission:
        commission[key] = round(commission[key], 3)

    with open(workdir_cft.replace('/', '\\') + '\\' + 'пины-комиссии из цфт.txt', 'w', encoding='utf-8') as file:
        for key in commission:
            file.write(str(key) + '\t' + str(round(commission[key],2)) + '\n')


def making_list_tables():
    xl = openpyxl.open(workdir_incas.replace('/', '\\') + '\\' + name_incas, read_only=True)

    names = xl.sheetnames

    sheets = []

    for name in names:
        sheets.append(xl[name])

    PINs = []
    commissions = []

    for sheet in sheets:
        row = 3
        while sheet[row][2].value != None:
            PINs.append(sheet[row][2].value)
            commissions.append(sheet[row][14].value)
            row += 1

    all_PINs = list(set(PINs))
    final = {}

    for each in all_PINs:
        final[each] = 0

    for PIN in all_PINs:
        for i in range(len(PINs)):
            if PIN == PINs[i]:
                final[PIN] += commissions[i]

    with open(workdir_incas + '\\пины-комиссии из сверок.txt', 'w', encoding="utf-8") as file:
        for key in final:
            file.write(str(key) + '\t' + str(round(final[key], 2)) + '\n')

def error():
    making_list_tables()
    making_list_cft()

    with open(workdir_incas + '\\пины-комиссии из сверок.txt', 'r', encoding='utf-8') as file:
        checking = file.read()

    with open(workdir_cft.replace('/', '\\') + '\\' + 'пины-комиссии из цфт.txt', 'r', encoding='utf-8') as file:
        cft = file.read()

    checking = checking.split('\n')
    cft = cft.split('\n')

    cft_lines = []
    checking_lines = []

    errors = []

    for line in checking:
        line = line.split('\t')
        checking_lines.append(line)

    for line in cft:
        line = line.split('\t')
        cft_lines.append(line)

    checking_lines = checking_lines[2:-2]
    cft_lines = cft_lines[2:-2]

    for each in checking_lines:
        for one in cft_lines:
            if each[0] == one[0]:
                if each[1] != one[1]:
                    errors.append(one[0] + '\t' + one[1])

    with open(workdir_cft.replace('/', '\\') + '\\' + 'ошибки.txt', 'w', encoding='utf-8') as file:
        for each in errors:
            file.write(each + '\n')

lw_files_cft.itemDoubleClicked.connect(event_cft)
lw_files_incas.itemDoubleClicked.connect(event_incas)
btn_dir_incas.clicked.connect(showFilenamesList_incas)
btn_dir_cft.clicked.connect(showFilenamesList_cft)
btn_instruction.clicked.connect(popup)
btn_check.clicked.connect(error)

app.exec()
