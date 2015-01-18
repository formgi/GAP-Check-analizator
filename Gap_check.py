# -*- coding: utf-8 -*-

# ===================================
__software__ = 'GAP-Check analyzer'
__author__ = 'Michael Belyansky'
__copyright__ = 'Metro Cash and Carry Ukraine, 2015' + chr(169)
__license__ = 'GNU GPL v3'
__version__ = "1.0.1"
__maintainer__ = 'Michael Belyansky'
__email__ = 'formgi@gmail.com'
__status__ = 'Production'
# ===================================

from PyQt4 import QtGui
from PyQt4 import QtCore

import time
import gap_check_rc
import getpass
import io
import os
import sys

# ===================================
def check_load(spl_scr):
    for i in range(0, 34):
        for a in range(0, 6):
            spl_scr.label.setText('<font size="2" color="grey">' + spl_scr.temp_txt_1 + '</font>')
            spl_scr.prog_spl.setValue(i)
            QtGui.qApp.processEvents()
            time.sleep(0.005)
    if not os.path.exists(window.dir_report.text().strip()):
        try:
            os.makedirs(window.dir_report.text().strip())
        except PermissionError:
            pass
    for i in range(35, 69):
        for a in range(0, 6):
            spl_scr.label.setText('<font size="2" color="grey">' + spl_scr.temp_txt_2 + '</font>')
            spl_scr.prog_spl.setValue(i)
            QtGui.qApp.processEvents()
            time.sleep(0.005)
    if not os.path.exists(window.xls_dir_report.text().strip()):
        try:
            os.makedirs(window.xls_dir_report.text().strip())
        except PermissionError:
            pass
    for i in range(70, 101):
        for a in range(0, 6):
            spl_scr.label.setText('<font size="2" color="grey">' + spl_scr.temp_txt_3 + '</font>')
            spl_scr.prog_spl.setValue(i)
            QtGui.qApp.processEvents()
            time.sleep(0.005)
    QtGui.qApp.processEvents()
    time.sleep(0.7)

# ===================================
class load_splash(QtGui.QSplashScreen):
    def __init__(self, parent=None):
        QtGui.QSplashScreen.__init__(self, parent)
        self.setWindowTitle('GAP-Check analyzer')
        self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(r':/icons/main.png')))
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.spl_pxm = QtGui.QPixmap(r':/icons/start_screen.png')
        self.setPixmap(self.spl_pxm)
        self.setMask(self.spl_pxm.mask())
        self.l_version = QtGui.QLabel(self)
        self.l_version.setGeometry(QtCore.QRect(205, 15, 90, 15))
        self.l_version.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignRight|QtCore.Qt.AlignVCenter)
        self.l_version.setText('<font size="2" color=#F7F30A>' + self.tr('version: ') + __version__ + '</font>')
        self.label = QtGui.QLabel(self)
        self.temp_txt_1 = self.tr('check load option')
        self.temp_txt_2 = self.tr('check directory report')
        self.temp_txt_3 = self.tr('check excel export directory')
        self.label.setGeometry(QtCore.QRect(20, 195, 255, 21))
        self.label.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.prog_spl = QtGui.QProgressBar(self)
        self.prog_spl.setTextVisible(False)
        style_pr_spl = '''QProgressBar{background-color: #323232; border: 1px solid blue; border-radius: 2px;}
                          QProgressBar::chunk{background-color: blue; background: yellow; width: 2px; margin: 0.5px }'''
        self.prog_spl.setStyleSheet(style_pr_spl)
        self.prog_spl.setGeometry(QtCore.QRect(20, 219, 278, 8))
        self.prog_spl.hide()
        self.show()
        for i in range(1, 101):
            self.setWindowOpacity(0.01 * i)
            QtGui.qApp.processEvents()
            time.sleep(0.01)
        self.label.show()
        self.prog_spl.show()
        check_load(self)
        for i in range(1, 101):
            self.setWindowOpacity(1 - 0.01 * i)
            QtGui.qApp.processEvents()
            time.sleep(0.01)
        self.finish(window)

# ===================================
def take_header():
    if window.setEnglish.isChecked() == True:
        header = ['№\narticle', 'Depart.', 'MU', 'Unit', 'Description']
    if window.setRussian.isChecked() == True:
        header = ['№\nарт.', 'Отдел', 'МЕ', 'Ед.\nизм.', 'Наименование']
    if window.setUkrainian.isChecked() == True:
        header = ['№\nарт.', 'Відділ', 'МО', 'Од.\nвим.', 'Найменування']
    return header

# ===================================
class make_table_model(QtCore.QAbstractTableModel):
    def __init__(self, datain, headerdata, parent = None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.arraydata = datain
        self.headerdata = headerdata
    def rowCount(self, parent):
        return len(self.arraydata)
    def columnCount(self, parent):
        return len(self.arraydata[0])
    def data(self, index, role):
        if not index.isValid():
            return None
        elif role == QtCore.Qt.TextAlignmentRole:
            if index.column() < 4:
                return QtCore.Qt.AlignCenter
            else:
                return QtCore.Qt.AlignLeft
        elif role != QtCore.Qt.DisplayRole:
            return None
        return self.arraydata[index.row()][index.column()]
    def headerData(self, col, orientation, role):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return self.headerdata[col]
            else:
                return col + 1
        return None
    def sort(self, col, order):
        if not col in [0, 1, 4]:
            window.tv.sortByColumn(0, QtCore.Qt.AscendingOrder)
            return None
        self.emit(QtCore.SIGNAL('layoutAboutToBeChanged()'))
        self.arraydata = sorted(self.arraydata, key = lambda num_col: num_col[col])
        if order == QtCore. Qt.DescendingOrder:
            self.arraydata.reverse()
        self.emit(QtCore.SIGNAL('layoutChanged()'))
        self.reset()

# ===================================
def make_table(art_data):
    window.tv = QtGui.QTableView(window)
    window.tab_mod = make_table_model(art_data, take_header())
    window.tab_mod.sort(0, QtCore.Qt.AscendingOrder)
    window.proxy = QtGui.QSortFilterProxyModel()
    window.proxy.setSourceModel(window.tab_mod)
    window.tv.setModel(window.proxy)
    window.tv.setGeometry(QtCore.QRect(10, 80, 550, 215))
    font = QtGui.QFont('Times New Roman', 10)
    window.tv.setFont(font)
    window.tv.setFrameShape(QtGui.QFrame.StyledPanel)
    window.tv.setFrameShadow(QtGui.QFrame.Sunken)
    window.tv.setLineWidth(1)
    window.tv.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
    window.tv.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
    horiz_head = window.tv.horizontalHeader()
    horiz_head.resizeSection(0, 75)
    horiz_head.resizeSection(1, 75)
    horiz_head.resizeSection(2, 55)
    horiz_head.resizeSection(3, 55)
    horiz_head.setResizeMode(0, 2)
    horiz_head.setResizeMode(1, 2)
    horiz_head.setResizeMode(2, 2)
    horiz_head.setResizeMode(3, 2)
    horiz_head.setResizeMode(4, 3)
    horiz_head.setFont(QtGui.QFont('Times New Roman', 10))
    window.tv.verticalHeader().setFont(QtGui.QFont('Times New Roman', 8))
    window.tv.verticalHeader().setDefaultSectionSize(20)
    window.tv.horizontalHeader().setStretchLastSection(True)
    window.tv.setSortingEnabled(True)
    window.tv.sortByColumn(0, QtCore.Qt.AscendingOrder)
    window.tv.setAlternatingRowColors(True)
    window.tv.show()

# ===================================
def take_space(n, stroka):
    rez = ''
    for i in stroka[n:]:
        if i != ' ':
            rez += i
        else:
            return rez

# ===================================
def take_alpha(n, stroka):
    rez = ''
    for i in stroka[n:]:
        if i.isalpha() == False:
            rez += i
        else:
            return rez

# ===================================
def format_row(row):
    import re
    filter_row = list()
    # article
    art_no = row[0:6]
    filter_row.append(art_no)
    # department
    dep = row[-4:].strip()
    filter_row.append(dep)
    # me & ed
    temp_ed = take_space(9, row)
    temp_me = take_alpha(0, temp_ed)
    if temp_me == '1,0':
        me = '1'
    else:
        me = temp_me
    filter_row.append(me)
    ed = row[len(temp_me)+9:len(temp_me)+11].lower()
    filter_row.append(ed)
    # description
    desc = row[len(temp_me)+11:re.search(r'\d?\d?\d?\d,\d\d', row).span()[0]].strip().rstrip(' .').strip('"').capitalize()
    filter_row.append(desc)
    return filter_row

# ===================================
def check_art(art_no, find_list):
    maxVal = window.progBarWin.progBar[0].maximum()
    check = None
    for i in range(5):
        QtGui.qApp.processEvents()
        window.progBarWin.progBar[1].setValue(i * 25)
        temp_list = find_list[i].getvalue()
        pos = temp_list.find(art_no)
        if pos > 0 and ord(temp_list[pos-1]) == 32:
            return art_no
    return check

# ===================================
def make_flist():
    find_list = list()
    for i in range(5):
        s = io.StringIO()
        f = open(window.lineEdit[i].text().rstrip())
        for line in f:
            if len(line) >= 3:
                s.write(line)
        find_list.append(s)
    return find_list

# ===================================
def make_list():
    s_list = io.StringIO()
    check_line = 0
    art_qty = 0
    tempstrfalse = ''
    f = open(window.lineEdit[5].text().rstrip())
    for line in f:
        code = ord(line[0])
        if code == 12:
            check_line = 1
        if check_line > 4 and len(line) > 1:
            tempstr = line.rstrip('\n')
            if line[0:6].isdigit() == True:
                if tempstrfalse != '':
                    art_qty += 1
                    s_list.write(tempstrfalse + '\n')
                    tempstrfalse = ''
                if len(tempstr) > 60:
                    s_list.write(tempstr + '\n')
                    art_qty += 1
                else:
                    tempstrfalse += tempstr
            else:
                tempstrfalse += tempstr
        if check_line > 0:
            check_line += 1
    s_list.seek(0)
    return s_list

# ===================================
def make_filter(art_data):
    if window.setEnglish.isChecked() == True:
        txt_filter = 'All'
    if window.setRussian.isChecked() == True:
        txt_filter = 'Все'
    if window.setUkrainian.isChecked() == True:
        txt_filter = 'Всі'
    window.filter_comboBox.addItem(txt_filter, 'All')
    dep_list = [art_data[row][1] for row in range(len(art_data))]
    dep_list_uniq = sorted(list(set(dep_list)))
    for i in range(len(dep_list_uniq)):
        window.filter_comboBox.addItem(dep_list_uniq[i], dep_list_uniq[i])

# ===================================
def start_report():
    window.filter_list = list()
    window.progBarWin.show()
    pos = window.pos()
    window.progBarWin.move(pos.x()+20, pos.y()+120)
    time.sleep(0.5)
    window.groupBox.hide()
    art_list = make_list()
    find_list = make_flist()
    window.progBarWin.progBar[0].setRange(0, len(art_list.readlines()) - 1)
    art_list.seek(0)
    for i, line in enumerate(art_list):
        form_line = format_row(line)
        QtGui.qApp.processEvents()
        maxVal = window.progBarWin.progBar[0].maximum()
        window.progBarWin.progBar[0].setValue(i + (maxVal - i) / 100)
        if i%50 == 0:
            window.progBarWin.label_pg[0].setText(form_line[0] + ' - ' + form_line[4])
        time.sleep(0.005)
        if form_line[1] in window.excep_dep.text():
            continue
        check = check_art(str(int(line.rstrip()[0:6])), find_list)
        if check == None:
            window.filter_list.append(form_line)
    window.progBarWin.hide()
    make_table(window.filter_list)
    make_filter(window.filter_list)
    window.report_button.hide()
    window.setFixedSize(570, 360)
    window.label_10.show()
    window.filter_comboBox.show()
    window.print_button.show()
    window.print_preview.show()
    window.export_xls.show()
    window.exit.show()

# ===================================
def take_option(mode, lang=0):
    # -----------------------------------
    link = os.getenv('APPDATA')
    s = QtCore.QSettings(link + r'\Metro Software\Gap-check\gap_check.ini',
        QtCore.QSettings.IniFormat)
    s.setIniCodec("UTF-8")
    # -----------------------------------
    if mode == 'load':
        window.dir_report.setText(s.value('Base/dir_report'))
        if window.dir_report.text() == '':
            window.dir_report.setText(r'C:\Temp\Gap_check')
            s.setValue('Base/dir_report', window.dir_report.text())
        # -----------------------------------
        window.xls_dir_report.setText(s.value('Base/xls_report'))
        if window.xls_dir_report.text() == '':
            window.xls_dir_report.setText(r'C:\Temp\Gap_check\Xls_reports')
            s.setValue('Base/xls_report', window.xls_dir_report.text())
        # -----------------------------------
        window.excep_dep.setText(s.value('Base/dep_exception'))
        if window.excep_dep.text() == '':
           window.excep_dep.setText('27,28')
           s.setValue('Base/dep_exception', window.excep_dep.text())
        window.excep_dep.textChanged.emit(window.excep_dep.text())
        if lang == 1:
            # -----------------------------------
            lng_int = s.value('Language/lng_int')
            if lng_int == None:
                lng_int = 'English'
                s.setValue('Language/lng_int', lng_int)
            if lng_int == 'English':
                window.setEnglish.setChecked(True)
            elif lng_int == 'Russian':
                window.setRussian.setChecked(True)
            else:
                window.setUkrainian.setChecked(True)
    elif mode == 'save':
        s.setValue('Base/dir_report', window.dir_report.text())
        s.setValue('Base/xls_report', window.xls_dir_report.text())
        s.setValue('Base/dep_exception', window.excep_dep.text())
        if window.setEnglish.isChecked() == True:
            s.setValue('Language/lng_int', 'English')
        if window.setRussian.isChecked() == True:
            s.setValue('Language/lng_int', 'Russian')
        if window.setUkrainian.isChecked() == True:
            s.setValue('Language/lng_int', 'Ukrainian')

# ===================================
class set_option_win(QtGui.QWidget):
    def __init__(self):
        super(set_option_win, self).__init__()
    def closeEvent(self, event):
        take_option('load')

# ===================================
class setup_win(QtGui.QMainWindow):
    def __init__(self, parent=None):
        # ===================================
        QtGui.QMainWindow.__init__(self, parent)
        self.setFixedSize(570, 300)
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/main.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.setWindowFlags(QtCore.Qt.MSWindowsFixedSizeDialogHint)
        QtGui.QApplication.setStyle(QtGui.QStyleFactory.create('Fusion'))
        QtGui.QApplication.setPalette(QtGui.QApplication.palette())

        # ===================================
        self.dateEdit = QtGui.QDateEdit(self)
        self.dateEdit.setGeometry(QtCore.QRect(160, 30, 100, 21))
        self.dateEdit.setFocusPolicy(QtCore.Qt.NoFocus)
        self.dateEdit.setDate(QtCore.QDate.currentDate())
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        self.dateEdit.setFont(font)

        # ===================================
        self.label = QtGui.QLabel(self)
        self.label.setGeometry(QtCore.QRect(10, 30, 150, 21))
        font = QtGui.QFont()
        font.setBold(True)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)

        # ===================================
        self.groupBox = QtGui.QGroupBox(self)
        self.groupBox.setGeometry(QtCore.QRect(10, 60, 550, 210))
        font = QtGui.QFont()
        font.setBold(True)
        self.groupBox.setFont(font)

        # ===================================
        self.lineEdit = list()
        self.label_le = list()
        set_y = 30
        for i in range(6):
            self.lineEdit.append(QtGui.QLineEdit(self.groupBox))
            self.lineEdit[i].setGeometry(QtCore.QRect(210, set_y, 290, 21))
            self.lineEdit[i].setReadOnly(True)
            self.label_le.append(QtGui.QLabel(self.lineEdit[i]))
            self.label_le[i].setStyleSheet('QLabel{background-color : white}')
            self.label_le[i].setPixmap(QtGui.QPixmap(r':/icons/cross.png'))
            self.label_le[i].setScaledContents(True)
            layout = QtGui.QHBoxLayout(self.lineEdit[i])
            layout.addWidget(self.label_le[i],0,QtCore.Qt.AlignVCenter|QtCore.Qt.AlignRight)
            layout.setSpacing(0)
            layout.setMargin(3)
            set_y += 30

        # ===================================
        mapper = QtCore.QSignalMapper(self)
        set_y = 30
        for i in range(6):
            self.toolButton = QtGui.QToolButton(self.groupBox)
            self.toolButton.setText('...')
            self.toolButton.setGeometry(QtCore.QRect(510, set_y, 25, 21))
            self.connect(self.toolButton, QtCore.SIGNAL("clicked()"), mapper, QtCore.SLOT("map()"))
            mapper.setMapping(self.toolButton, i)
            set_y += 30
        self.connect(mapper, QtCore.SIGNAL("mapped(int)"), self.file_name)

        # ===================================
        set_y = 30
        self.lfile = list()
        for i in range(6):
            self.lfile.append(QtGui.QLabel(self.groupBox))
            self.lfile[i].setGeometry(QtCore.QRect(20, set_y, 191, 21))
            set_y += 30

        # ===================================
        self.label_8 = QtGui.QLabel(self)
        self.label_8.setGeometry(QtCore.QRect(280, 30, 191, 21))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        self.label_8.setPalette(palette)
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setFrameShadow(QtGui.QFrame.Plain)
        self.label_8.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)

        # ===================================
        self.label_9 = QtGui.QLabel(self)
        self.label_9.setGeometry(QtCore.QRect(490, 30, 71, 21))
        self.label_9.setPixmap(QtGui.QPixmap(r':/icons/metro.png'))
        self.label_9.setScaledContents(True)

        # ===================================
        self.report_button = QtGui.QPushButton(self)
        self.report_button.setGeometry(QtCore.QRect(10, 285, 120, 25))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/start_report.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.report_button.setIcon(icon)
        self.report_button.setIconSize(QtCore.QSize(16, 16))
        QtCore.QObject.connect(self.report_button, QtCore.SIGNAL('clicked()'), self.push_button)
        self.report_button.hide()

        # ===================================
        self.label_10 = QtGui.QLabel(self)
        self.label_10.setGeometry(QtCore.QRect(10, 55, 165, 21))
        font = QtGui.QFont()
        font.setBold(True)
        self.label_10.setFont(font)
        # ===================================
        self.filter_comboBox = QtGui.QComboBox(self)
        self.filter_comboBox.setGeometry(QtCore.QRect(205, 55, 55, 21))
        QtCore.QObject.connect(self.filter_comboBox, QtCore.SIGNAL('activated(int)'), self.set_filter)
        # ===================================
        self.print_button = QtGui.QPushButton(self)
        self.print_button.setGeometry(QtCore.QRect(10, 310, 120, 25))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/printer.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.print_button.setIcon(icon)
        self.print_button.setIconSize(QtCore.QSize(16, 16))
        QtCore.QObject.connect(self.print_button, QtCore.SIGNAL('clicked()'), self.printForm)
        # ===================================
        self.print_preview = QtGui.QPushButton(self)
        self.print_preview.setGeometry(QtCore.QRect(140, 310, 120, 25))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/preview.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.print_preview.setIcon(icon)
        self.print_preview.setIconSize(QtCore.QSize(16, 16))
        QtCore.QObject.connect(self.print_preview, QtCore.SIGNAL('clicked()'), self.PrintPreview)
        # ===================================
        self.export_xls = QtGui.QPushButton(self)
        self.export_xls.setGeometry(QtCore.QRect(270, 310, 180, 25))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/excel.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.export_xls.setIcon(icon)
        self.export_xls.setIconSize(QtCore.QSize(16, 16))
        QtCore.QObject.connect(self.export_xls, QtCore.SIGNAL('clicked()'), self.exportXLS)

        # ===================================
        self.exit = QtGui.QPushButton(self)
        self.exit.setGeometry(QtCore.QRect(460, 310, 100, 25))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(r':/icons/exit.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.exit.setIcon(icon)
        self.exit.setIconSize(QtCore.QSize(16, 16))
        QtCore.QObject.connect(self.exit, QtCore.SIGNAL('clicked()'), self.close)

        # ===================================
        self.label_10.hide()
        self.filter_comboBox.hide()
        self.print_button.hide()
        self.print_preview.hide()
        self.export_xls.hide()
        self.exit.hide()

        # ===================================
        self.progBarWin = QtGui.QWidget(self)
        self.progBarWin.progBar = list()
        self.progBarWin.label_pg = list()
        layout = QtGui.QVBoxLayout()
        font = QtGui.QFont()
        font.setPointSize(10)
        for i in range(2):
            self.progBarWin.progBar.append(QtGui.QProgressBar())
            font.setBold(False)
            self.progBarWin.progBar[i].setFont(font)
            self.progBarWin.progBar[i].setMaximum(100)
            self.progBarWin.progBar[i].setTextVisible(True)
            self.progBarWin.progBar[i].setOrientation(QtCore.Qt.Horizontal)
            self.progBarWin.progBar[i].setTextDirection(QtGui.QProgressBar.TopToBottom)
            self.progBarWin.label_pg.append(QtGui.QLabel(self))
            self.progBarWin.label_pg[i].setAlignment(QtCore.Qt.AlignVCenter)
            font.setBold(True)
            self.progBarWin.label_pg[i].setFont(font)
            layout.addWidget(self.progBarWin.label_pg[i])
            layout.addWidget(self.progBarWin.progBar[i])
        self.progBarWin.label_pg[1].setText(self.tr('gap-check file:'))
        self.progBarWin.setLayout(layout)
        self.progBarWin.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.FramelessWindowHint)
        self.progBarWin.resize(530, 50)
        self.progBarWin.setWindowModality(QtCore.Qt.WindowModal)

        # ===================================
        self.optionWin = set_option_win()
        self.optionWin.resize(350, 250)
        self.optionWin.setWindowModality(QtCore.Qt.WindowModal)
        self.optionWin.setWindowFlags(QtCore.Qt.Drawer)
        # -----------------------------------
        self.label_1 = QtGui.QLabel(self.optionWin)
        self.label_1.setGeometry(QtCore.QRect(10, 10, 250, 21))
        # -----------------------------------
        self.dir_report = QtGui.QLineEdit(self.optionWin)
        self.dir_report.setGeometry(QtCore.QRect(10, 35, 250, 21))
        self.dir_report.setReadOnly(True)
        # -----------------------------------
        self.but_dir_rep = QtGui.QPushButton(self.optionWin)
        self.but_dir_rep.setGeometry(QtCore.QRect(270, 35, 70, 21))
        self.but_dir_rep.clicked.connect(self.take_dir_report)
        # -----------------------------------
        self.label_2 = QtGui.QLabel(self.optionWin)
        self.label_2.setGeometry(QtCore.QRect(10, 75, 250, 21))
        # -----------------------------------
        self.xls_dir_report = QtGui.QLineEdit(self.optionWin)
        self.xls_dir_report.setGeometry(QtCore.QRect(10, 100, 250, 21))
        self.xls_dir_report.setReadOnly(True)
        # -----------------------------------
        self.but_xls_dir_rep = QtGui.QPushButton(self.optionWin)
        self.but_xls_dir_rep.setGeometry(QtCore.QRect(270, 100, 70, 21))
        self.but_xls_dir_rep.clicked.connect(self.take_xlsdir_report)
        # -----------------------------------
        self.label_3 = QtGui.QLabel(self.optionWin)
        self.label_3.setGeometry(QtCore.QRect(10, 140, 300, 21))
        # -----------------------------------
        self.excep_dep = QtGui.QLineEdit(self.optionWin)
        self.excep_dep.setGeometry(QtCore.QRect(10, 165, 240, 21))
        regexp = QtCore.QRegExp('(\d\d\,){1,5}\d\d')
        validator = QtGui.QRegExpValidator(regexp)
        self.excep_dep.setValidator(validator)
        self.excep_dep.textChanged.connect(self.check_state)

        # -----------------------------------
        self.label_4 = QtGui.QLabel(self.optionWin)
        self.label_4.setGeometry(QtCore.QRect(260, 165, 100, 21))
        # -----------------------------------
        self.button_save = QtGui.QPushButton(self.optionWin)
        self.button_save.setIcon(QtGui.QIcon(r':/icons/tick.png'))
        self.button_save.setIconSize(QtCore.QSize(16, 16))
        self.button_save.setGeometry(QtCore.QRect(10, 210, 120, 25))
        self.button_save.clicked.connect(self.save_option)
        # -----------------------------------
        self.button_cancel = QtGui.QPushButton(self.optionWin)
        self.button_cancel.setIcon(QtGui.QIcon(r':/icons/cross.png'))
        self.button_cancel.setIconSize(QtCore.QSize(16, 16))
        self.button_cancel.setGeometry(QtCore.QRect(140, 210, 120, 25))
        self.button_cancel.clicked.connect(self.close_option)

        # ===================================
        self.xls_importAct = QtGui.QAction(self, shortcut='Ctrl+E', triggered=self.exportXLS)
        self.xls_importAct.setIcon(QtGui.QIcon(r':/icons/excel.png'))
        self.xls_importAct.setEnabled(False)
        # -----------------------------------
        self.printAct = QtGui.QAction(self, shortcut=QtGui.QKeySequence.Print, triggered=self.printForm)
        self.printAct.setIcon(QtGui.QIcon(r':/icons/printer.png'))
        self.printAct.setEnabled(False)
        # -----------------------------------
        self.printViewAct = QtGui.QAction(self, shortcut='Ctrl+W', triggered=self.PrintPreview)
        self.printViewAct.setIcon(QtGui.QIcon(r':/icons/preview.png'))
        self.printViewAct.setEnabled(False)
        # -----------------------------------
        self.exitAct = QtGui.QAction(self, shortcut='Ctrl+Q', triggered=self.close)
        self.exitAct.setIcon(QtGui.QIcon(r':/icons/exit.png'))
        # -----------------------------------
        self.optionsAct = QtGui.QAction(self, triggered=self.options)
        self.optionsAct.setIcon(QtGui.QIcon(r':/icons/options.png'))
        # -----------------------------------
        self.setEnglish = QtGui.QAction(QtGui.QIcon(r':/icons/en.png'), 'English', self,
                statusTip='Setup english interface', checkable=True,
                triggered=self.setInterface)
        self.setRussian = QtGui.QAction(QtGui.QIcon(r':/icons/ru.png'), 'Русский', self, checkable=True,
                statusTip='Установить русский интерфейс',
                triggered=self.setInterface)
        self.setUkrainian = QtGui.QAction(QtGui.QIcon(r':/icons/ua.png'), 'Українська', self, checkable=True,
                statusTip='Застосувати український інтерфейс',
                triggered=self.setInterface)
        # -----------------------------------
        self.aboutAct = QtGui.QAction(self, triggered=self.about)
        self.aboutAct.setIcon(QtGui.QIcon(r':/icons/info.png'))
        # -----------------------------------
        self.languageGroup = QtGui.QActionGroup(self)
        self.languageGroup.addAction(self.setEnglish)
        self.languageGroup.addAction(self.setRussian)
        self.languageGroup.addAction(self.setUkrainian)
        self.setEnglish.setChecked(True)
        # -----------------------------------
        self.fileMenu = self.menuBar().addMenu('')
        self.fileMenu.addAction(self.xls_importAct)
        self.fileMenu.addAction(self.printAct)
        self.fileMenu.addAction(self.printViewAct)
        self.fileMenu.addSeparator()
        self.fileMenu.addAction(self.exitAct)

        self.toolsMenu = self.menuBar().addMenu('')
        self.toolsMenu.addAction(self.optionsAct)
        self.languageMenu = self.toolsMenu.addMenu(QtGui.QIcon(r':/icons/lg.png'), '')
        self.languageMenu.addAction(self.setEnglish)
        self.languageMenu.addAction(self.setRussian)
        self.languageMenu.addAction(self.setUkrainian)

        self.helpMenu = self.menuBar().addMenu('')
        self.helpMenu.addAction(self.aboutAct)

        self.statusBar()

        self.setInterface()

    # ===================================
    def check_state(self):
        sender = self.sender()
        validator = sender.validator()
        state = validator.validate(sender.text(), 0)[0]
        if state == QtGui.QValidator.Acceptable:
            color = '#c4df9b' # green
        elif state == QtGui.QValidator.Intermediate:
            color = '#fff79a' # yellow
        else:
            color = '#f6989d' # red
        sender.setStyleSheet('QLineEdit {background-color: %s}' % color)

    # ===================================
    def save_option(self):
        take_option('save')
        self.optionWin.hide()

    # ===================================
    def close_option(self):
        take_option('load')
        self.optionWin.hide()

    # ===================================
    def take_dir_report(self):
        temp_dir = QtGui.QFileDialog.getExistingDirectory(parent=self.optionWin, directory=self.dir_report.text())
        if temp_dir:
            self.dir_report.setText(temp_dir.replace('/', '\\'))

    # ===================================
    def take_xlsdir_report(self):
        temp_dir = QtGui.QFileDialog.getExistingDirectory(parent=self.optionWin, directory=self.xls_dir_report.text())
        if temp_dir:
            self.xls_dir_report.setText(temp_dir.replace('/', '\\'))

   # ===================================
    def closeEvent(self, event):
        dialog = QtGui.QMessageBox(QtGui.QMessageBox.Question, self.tr('GAP-Check analyzer'), self.tr('You are sure to quit?'),
            buttons = QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, parent=self)
        dialog.button(QtGui.QMessageBox.Yes).setText(self.tr('&Yes'))
        dialog.button(QtGui.QMessageBox.Yes).setIcon(QtGui.QIcon(r':/icons/tick.png'))
        dialog.button(QtGui.QMessageBox.No).setText(self.tr('&No'))
        dialog.button(QtGui.QMessageBox.No).setIcon(QtGui.QIcon(r':/icons/cross.png'))
        reply = dialog.exec_()
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
            take_option('save')
        else:
            event.ignore()

    # ===================================
    def exportXLS(self):
        import xlsxwriter
        currentData = QtCore.QDateTime.currentDateTime().date().toString('dd.MM.yyyy')
        currentTime = QtCore.QDateTime.currentDateTime().time().toString('hh:mm:ss')
        filename = 'Gap_check_' + window.dateEdit.text() + '.xlsx'
        full_filename = self.xls_dir_report.text()+ '\\' + filename
        workbook = xlsxwriter.Workbook(full_filename)
        mod = self.tv.model()
        rows = mod.rowCount()
        columns = mod.columnCount()
        header = take_header()
        worksheet = workbook.add_worksheet('Gap_check')
        worksheet.set_header('&RGap Check report MCC UA')
        worksheet.set_footer('&RPage &P of &N')
        worksheet.set_column(0, 1, 15)
        worksheet.set_column(2, 3, 10)
        worksheet.set_column(4, 4, 32)
        header_txt = self.tr('report created ') + currentData + self.tr(' on ') + currentTime + self.tr(' user: ') + user_name
        worksheet.write('E1', header_txt, workbook.add_format({'font_size': 8, 'align': 'right'}))
        worksheet.set_row(0, 12)
        worksheet.merge_range('A2:E2', 'UNSCANNED ARTICLES at ' + window.dateEdit.text(), workbook.add_format({'bold': True, 'font_size':16, 'align': 'center'}))
        form_main = {'bold': True, 'font_name': 'arial', 'align': 'center', 'valign': 'vcenter'}
        form_h = list()
        for i in range(5):
            form_h.append(form_main.copy())
        form_h[0].update({'top': 2, 'bottom': 2, 'left': 2, 'right': 1})
        form_h[1].update({'top': 2, 'bottom': 2, 'left': 1, 'right': 1})
        form_h[2].update({'top': 2, 'bottom': 2, 'left': 1, 'right': 1})
        form_h[3].update({'top': 2, 'bottom': 2, 'left': 1, 'right': 1})
        form_h[4].update({'top': 2, 'bottom': 2, 'left': 1, 'right': 2})
        for i in range(5):
            worksheet.write(2, i, header[i], workbook.add_format(form_h[i]))
        worksheet.set_row(2, 30)
        for r in range(rows):
            for c in range(columns):
                ind = mod.index(r, c)
                value = mod.data(ind, QtCore.Qt.DisplayRole)
                if c < 3:
                    worksheet.write_number(r + 3, c, int(value), workbook.add_format({'align': 'center', 'border': 1}))
                else:
                    worksheet.write_string(r + 3, c, value,  workbook.add_format({'align': 'center', 'border': 1}) if c == 3 else workbook.add_format({'align': 'left', 'border': 1}))
        worksheet.autofilter('A3:E3')
        worksheet.repeat_rows(2)
        workbook.set_properties({'title': 'Gap-check report list', 'subject': 'Unscanned articles from Gap-check report', 'author': user_name,
                                 'company': 'Metro Cash and Carry Ukraine', 'category': 'Reports',
                                 'comments': 'Created with Python and XlsxWriter,\nprogram: Gap-check analizer'})
        try:
            workbook.close()
            msgInfo = self.tr('Workbook:<br>') + full_filename + self.tr('<br>save successfully!<br>Open this file?')
            dialog = QtGui.QMessageBox(QtGui.QMessageBox.Question, self.tr('Export to MS Excel'), msgInfo,
                        buttons = QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, parent=self)
            dialog.button(QtGui.QMessageBox.Yes).setText(self.tr('&Yes'))
            dialog.button(QtGui.QMessageBox.Yes).setIcon(QtGui.QIcon(r':/icons/tick.png'))
            dialog.button(QtGui.QMessageBox.No).setText(self.tr('&No'))
            dialog.button(QtGui.QMessageBox.No).setIcon(QtGui.QIcon(r':/icons/cross.png'))
            reply = dialog.exec_()
            if reply == QtGui.QMessageBox.Yes:
                os.startfile(full_filename)
        except PermissionError:
            msgError = self.tr('Error save workbook,<br>please close file:<br><br>') + filename
            QtGui.QMessageBox.warning(self, self.tr('Error save workbook'), msgError)

    # ===================================
    def push_button(self):
        self.report_button.setEnabled(False)
        self.groupBox.setEnabled(False)
        QtGui.qApp.processEvents()
        self.xls_importAct.setEnabled(True)
        self.printAct.setEnabled(True)
        self.printViewAct.setEnabled(True)
        time.sleep(1)
        start_report()

    # ===================================
    def printForm(self):
        printer = QtGui.QPrinter(QtGui.QPrinter.ScreenResolution)
        printer.setPaperSize(QtGui.QPrinter.A4)
        self.pdialog = dialog = QtGui.QPrintDialog(printer)
        dialog.setOption(QtGui.QAbstractPrintDialog.PrintShowPageSize, False)
        if dialog.exec_() == QtGui.QDialog.Accepted:
            printer = dialog.printer()
            self.paintPreview(printer)

    # ===================================
    def PrintPreview(self):
        printer = QtGui.QPrinter(QtGui.QPrinter.ScreenResolution)
        printer.setPaperSize(QtGui.QPrinter.A4)
        preview = QtGui.QPrintPreviewDialog(printer)
        preview.paintRequested.connect(self.paintPreview)
        preview.exec_()

    # ===================================
    def paintPreview(self, printer):
        # -----------------------------------
        def paint_header():
            header = take_header()
            head_rect = [[50, 140], [140, 210], [210, 260], [260, 350], [350, 750]]
            painter.setPen(QtGui.QPen(QtCore.Qt.black, 1, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
            painter.setFont(QtGui.QFont('Times New Roman', 10, QtGui.QFont.Bold))
            for i in range(5):
                text_rect = QtCore.QRect(QtCore.QPoint(head_rect[i][0], 40), QtCore.QPoint(head_rect[i][1], 70))
                painter.drawText(text_rect, QtCore.Qt.AlignCenter, header[i])
            painter.drawLine(50, 40, 750, 40)
            painter.drawLine(50, 70, 750, 70)
            painter.drawLine(50, 40, 50, 70)
            painter.drawLine(750, 40, 750, 70)
            painter.setPen(QtGui.QPen(QtCore.Qt.black, 0.5, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
            painter.drawLine(140, 40, 140, 70)
            painter.drawLine(210, 40, 210, 70)
            painter.drawLine(260, 40, 260, 70)
            painter.drawLine(350, 40, 350, 70)
        # -----------------------------------
        def paint_row(numb, mode):
            y_1 = 70 + numb * 25
            y_2 = 95 + numb * 25
            if mode == 0:
                painter.setPen(QtGui.QPen(QtCore.Qt.black, 1, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
                painter.drawLine(50, y_1, 50, y_2)
                painter.drawLine(750, y_1, 750, y_2)
                painter.setPen(QtGui.QPen(QtCore.Qt.black, 0.5, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
                painter.drawLine(140, y_1, 140, y_2)
                painter.drawLine(210, y_1, 210, y_2)
                painter.drawLine(260, y_1, 260, y_2)
                painter.drawLine(350, y_1, 350, y_2)
            elif mode == 1:
                painter.setPen(QtGui.QPen(QtCore.Qt.black, 1, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
                painter.drawLine(50, y_2, 750, y_2)
            else:
                painter.setPen(QtGui.QPen(QtCore.Qt.black, 0.5, QtCore.Qt.SolidLine, QtCore.Qt.RoundCap, QtCore.Qt.RoundJoin))
                painter.drawLine(50, y_2, 750, y_2)
        # -----------------------------------
        painter = QtGui.QPainter()
        painter.begin(printer)
        mod = self.tv.model()
        rows = mod.rowCount()

        orient = printer.orientation()
        page_range = printer.printRange()
        if page_range == 2:
            from_page = printer.fromPage()
            to_page = printer.toPage()
        else:
            from_page = 0
            to_page = 0
        if orient == 0:
            rows_per_page = 40
        else:
            rows_per_page = 26
        if page_range != 2 or from_page == 1:
            currentData = QtCore.QDateTime.currentDateTime().date().toString('dd.MM.yyyy')
            currentTime = QtCore.QDateTime.currentDateTime().time().toString('hh:mm:ss')
            header_txt = self.tr('report created ') + currentData + self.tr(' on ') + currentTime + self.tr(' user: ') + user_name
            header_top = self.tr('UNSCANNED ARTICLES at ') + window.dateEdit.text()
            painter.setFont(QtGui.QFont('Times New Roman', 16, QtGui.QFont.Bold))
            painter.drawText(QtCore.QRect(QtCore.QPoint(50, 5), QtCore.QPoint(750, 35)), QtCore.Qt.AlignCenter, header_top)
            painter.setFont(QtGui.QFont('Times New Roman', 7))
            fm = QtGui.QFontMetrics(QtGui.QFont('Times New Roman', 7))
            w = int(fm.width(header_txt))
            painter.drawText(750 - w, 37, header_txt)
        paint_header()
        pages = rows // rows_per_page
        if rows % rows_per_page != 0:
            last_page_rows = rows % rows_per_page - 1
            pages += 1
        if page_range == 2:
            pages_pr = to_page - from_page + 1
        else:
            pages_pr = pages - 1
        progress = QtGui.QProgressDialog(self.tr('make report...'), self.tr('Abort print'), 1, pages_pr, self)
        progress.setWindowTitle(self.tr('Print report'))
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowTitleHint)
        page = 1
        progress.show()
        progress.setValue(page)
        QtGui.qApp.processEvents()
        time.sleep(0.5)
        row = 0
        y_top = 70
        y_bot = 95
        row_x = [[50, 140], [140, 210], [210, 260], [260, 350], [360, 750]]
        for r in range(rows):
            if progress.wasCanceled():
                break
            if page_range == 2 and page > to_page:
                continue
            paint_row(row, 0)
            row_rect = list()
            for i in range(5):
                row_rect.append(QtCore.QRect(QtCore.QPoint(row_x[i][0], y_top), QtCore.QPoint(row_x[i][1], y_bot)))
            for c in range(5):
                ind = mod.index(r, c)
                if c < 4:
                    align = QtCore.Qt.AlignCenter
                else:
                    align = QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
                painter.setFont(QtGui.QFont('Arial', 10))
                painter.drawText(row_rect[c], align, mod.data(ind, QtCore.Qt.DisplayRole))
            if row == rows_per_page - 1:
                paint_row(row, 1)
                bottom_txt = self.tr('Page ') + str(page) + self.tr(' of  ') + str(pages)
                painter.setFont(QtGui.QFont('Times New Roman', 7))
                fm = QtGui.QFontMetrics(QtGui.QFont('Times New Roman', 7))
                w = int(fm.width(bottom_txt))
                h = int(fm.height())
                painter.drawText(750 - w, (rows_per_page * 25 + 75) + h, bottom_txt)
            else:
                if page == pages and row == last_page_rows:
                    paint_row(row, 1)
                    bottom_txt = self.tr('Page ') + str(page) + self.tr(' of  ') + str(pages)
                    painter.setFont(QtGui.QFont('Times New Roman', 7))
                    w = int(fm.width(bottom_txt))
                    h = int(fm.height())
                    painter.drawText(750 - w, (rows_per_page * 25 + 75) + h, bottom_txt)
                else:
                    paint_row(row, 2)
            row += 1
            y_top += 25
            y_bot += 25
            if row == rows_per_page:
                if page_range == 2:
                    if page >= from_page and page < to_page:
                        progress.setValue(page)
                        time.sleep(0.5)
                        printer.newPage()
                    elif page == to_page:
                        break
                    else:
                        painter.eraseRect(0, 0, 800, rows_per_page * 25 + 200)
                else:
                    progress.setValue(page)
                    time.sleep(0.5)
                    printer.newPage()
                paint_header()
                page += 1
                row = 0
                y_top = 70
                y_bot = 95
        painter.end()

    # ===================================
    def options(self):
        take_option('load')
        self.optionWin.show()

    # ===================================
    def setInterface(self):
        translator.load('')
        QtGui.qApp.installTranslator(translator)
        if self.setRussian.isChecked() == True:
            translator.load(r':/translate/gap_check_ru.qm')
            QtGui.qApp.installTranslator(translator)
        if self.setUkrainian.isChecked() == True:
            translator.load(r':/translate/gap_check_ua.qm')
            QtGui.qApp.installTranslator(translator)
        # -----------------------------------
        self.setWindowTitle('GAP-Check analyzer')
        self.label.setText(self.tr('Analysis GAP-Check on:'))
        self.groupBox.setTitle(self.tr('Select text files for analysis:'))
        self.lfile[0].setText(self.tr('Non/Late delivery (1/2)'))
        self.lfile[1].setText(self.tr('Not ordered Store (3)'))
        self.lfile[2].setText(self.tr('Article not present (4)'))
        self.lfile[3].setText(self.tr('Not ord. HO/Art. blok (5/6)'))
        self.lfile[4].setText(self.tr('Deleted article (7)'))
        self.lfile[5].setText(self.tr('Availability Stock <=0'))
        self.label_8.setText('GAP-Check analyzer')
        self.report_button.setText(self.tr('&Start report'))
        self.label_10.setText(self.tr('Set filter on departmen(s):'))
        self.print_button.setText(self.tr('&Print report'))
        self.print_preview.setText(self.tr('Print previe&w'))
        self.export_xls.setText(self.tr('Export to Excel 2010'))
        self.exit.setText(self.tr('E&xit'))
        # -----------------------------------
        self.xls_importAct.setText(self.tr('Expo&rt to Excel'))
        self.xls_importAct.setStatusTip(self.tr('Export to Excel 2010'))
        self.printAct.setText(self.tr('&Print'))
        self.printAct.setStatusTip(self.tr('Print the table'))
        self.printViewAct.setText(self.tr('Print previe&w'))
        self.printViewAct.setStatusTip(self.tr('Preview table for print'))
        self.exitAct.setText(self.tr('E&xit'))
        self.exitAct.setStatusTip(self.tr('Exit the application'))
        self.optionsAct.setText(self.tr('O&ptions'))
        self.optionsAct.setStatusTip(self.tr('Options'))
        self.aboutAct.setText(self.tr('&About'))
        self.aboutAct.setStatusTip(self.tr('About the program'))
        self.fileMenu.setTitle(self.tr('&File'))
        self.toolsMenu.setTitle(self.tr('&Tools'))
        self.languageMenu.setTitle(self.tr('L&anguage'))
        self.helpMenu.setTitle(self.tr('&Help'))
        # -----------------------------------
        self.progBarWin.progBar[0].setFormat(self.tr('check article file %p%'))
        self.progBarWin.label_pg[1].setText(self.tr('gap-check file:'))
        # -----------------------------------
        self.optionWin.setWindowTitle(self.tr('Options'))
        self.label_1.setText(self.tr('Directory for report:'))
        self.but_dir_rep.setText(self.tr('Browse'))
        self.label_2.setText(self.tr('Directory for export xls:'))
        self.but_xls_dir_rep.setText(self.tr('Browse'))
        self.label_3.setText(self.tr('Exception department(s) in report:'))
        self.label_4.setText(self.tr('(example: 12,13)'))
        self.button_save.setText(self.tr('&Save'))
        self.button_cancel.setText(self.tr('&Cancel'))
        if hasattr(self, 'filter_list'):
            self.filter_comboBox.clear()
            make_filter(self.filter_list)
            self.tab_mod = make_table_model(self.filter_list, take_header())
            self.tab_mod.sort(0, QtCore.Qt.AscendingOrder)
            self.proxy.setSourceModel(self.tab_mod)
            self.tv.setModel(self.proxy)

    # ===================================
    def about(self):
        about_txt = '<HTML>'\
            '<p style="font-size:16px; color:#000099"><b>GAP-Check Analyzer</b></p>'\
            '<b>This program is provided to search unscanned</b><br>'\
            '<b>articles in reports the procedure GAP-Check</b><br>'\
            '<b>---------------------------------------------------------</b><br>' +\
            __copyright__ + '<br>'\
            'author: ' + __author__ + '<br>'\
            'e-mail: <a href="mailto:' + __email__ + '?subject=GAP-Check analizator'\
            ' software &body=Please send me a copy of your new program!">' + __email__ +'</a><br>'\
            'licence: ' + '<a href="LICENSE.txt">' + __license__ + '</a><br>'\
            'source: ' + '<a href="https://github.com/formgi/GAP-Check-analizator.git">Github repositories</a><br>'\
            'version: '+ __version__ + ' <br>'\
            '<b>---------------------------------------------------------</b><br>'\
            'GAP-Check analyzer developed in Python v. 3.4<br>'\
            '<a href="https://www.python.org/">https://www.python.org/</a><br><br>'\
            '<u><b>Used components and resources:</b></u><br>'\
            'PyQt GPL v4.11.2 for Python v3.4 (x32):<br>'\
            '<a href="www.riverbankcomputing.com/">www.riverbankcomputing.com/</a><br>'\
            'XlsxWriter is module for import to XLSX file format:<br>'\
            '<a href="https://xlsxwriter.readthedocs.org/">https://xlsxwriter.readthedocs.org</a><br>'\
            'Compile with cx_Freeze module:<br>'\
            '<a href="http://cx-freeze.readthedocs.org/">http://cx-freeze.readthedocs.org</a><br>'\
            'Silk icon set 1.3 by Mark James:<br>'\
            '<a href="http://www.famfamfam.com/lab/icons/silk/">http://www.famfamfam.com/lab/icons/silk</a><br>'\
            '<b>---------------------------------------------------------</b><br>'\
            '<center><b>Please <cite style="color:red">DONATE</cite> to<br>"GAP-Check analyzer" with PayPal:</b><br><br>'\
            '<a href="https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&amp;'\
            'hosted_button_id=YL72TBSVTCV3C"><img '\
            'src=":/icons/donate.png"<img/></a><br></center>'\
            '</HTML>'
        about_box = QtGui.QMessageBox(self)
        about_box.setWindowTitle(self.tr('About GAP-Check analyzer'))
        about_box.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(r':/icons/info.png')))
        about_box.setTextFormat(1)
        about_box.setText(about_txt)
        about_box.exec_()

    # ===================================
    def file_changed(self):
        empty = False
        for i in range(6):
            if bool(self.lineEdit[i].text()) == False:
                empty += False
            else:
                empty += True
        if empty == 6:
            time.sleep(0.5)
            self.setFixedSize(570, 335)
            self.report_button.show()
        else:
            if self.report_button.isVisible() == True:
                self.report_button.hide()
                self.setFixedSize(570, 300)

    # ===================================
    def file_name(self, i):
        name_report = [
            'Automatic Root Cause Не/Пізня доставка(1/2)',
            'Automatic Root Cause Недост. зам. ТЦ(3)',
            'Automatic Root Cause Артикул не викладен(4)',
            'Automatic Root Cause Недост. зам. ГО/Apт. заблок(5/6)',
            'Automatic Root Cause Артикул видалений(7)',
            'Назва звіту: Artikelliste']
        if self.lineEdit[i]:
            tempfiledb = self.lineEdit[i].text()
        filedb = (QtGui.QFileDialog.getOpenFileName(self, self.tr('Open file: ') + self.lfile[i].text(),
            self.dir_report.text(), 'Text files (*.txt)'))
        if not filedb:
            if tempfiledb:
                filedb = tempfiledb
                return
            else:
                return
        f = open(filedb, 'r')
        tempf = f.readlines()
        # check correct file
        check = True
        if len(tempf) > 24:
            data_form_15 = QtCore.QRegExp('Дата \d\d.\d\d.\d\d\d\d')
            data_form_6 = QtCore.QRegExp('\d\d.\d\d.\d\d\d\d')
            data_str_15 = tempf[21].rstrip()
            data_str_6 = tempf[2][0:10]
            if i < 5:
                name_str = tempf[24].rstrip()
            else:
                name_str = tempf[13].rstrip()
            val_data_15 = QtGui.QRegExpValidator(data_form_15)
            val_data_6 = QtGui.QRegExpValidator(data_form_6)
            valid_15 = val_data_15.validate(data_str_15, 0)[0]
            valid_6 = val_data_6.validate(data_str_6, 0)[0]
            if valid_15 == QtGui.QValidator.Acceptable or valid_6 == QtGui.QValidator.Acceptable:
                if valid_15 == QtGui.QValidator.Acceptable:
                    data_tmp = data_str_15[5:]
                if valid_6 == QtGui.QValidator.Acceptable:
                    data_tmp = data_str_6
            else:
                check = False
        else:
            check = False
        if check == False:
            QtGui.QMessageBox.warning(self, self.tr('Incorrect file'),
                self.tr('Wrong report file!<br>Unsupported format in file: <br><br>') + filedb)
            self.label_le[i].setPixmap(QtGui.QPixmap(':/icons/cross.png'))
            self.lineEdit[i].clear()
            self.file_changed()
            return
        # check date report
        if self.dateEdit.text() != data_tmp:
            msgError = self.tr('Data in file report: <br>') + data_tmp + self.tr('<br>Set this date?')
            dialog = QtGui.QMessageBox(QtGui.QMessageBox.Question, self.tr('Set data'), msgError,
                        buttons = QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, parent=self)
            dialog.button(QtGui.QMessageBox.Yes).setText(self.tr('&Yes'))
            dialog.button(QtGui.QMessageBox.Yes).setIcon(QtGui.QIcon(r':/icons/tick.png'))
            dialog.button(QtGui.QMessageBox.No).setText(self.tr('&No'))
            dialog.button(QtGui.QMessageBox.No).setIcon(QtGui.QIcon(r':/icons/cross.png'))
            reply = dialog.exec_()
            if reply == QtGui.QMessageBox.Yes:
                year = int(data_tmp[6:10])
                month = int(data_tmp[3:5])
                day = int(data_tmp[0:2])
                self.dateEdit.setDate(QtCore.QDate(year, month, day))
                for n in range(6):
                    self.lineEdit[n].clear()
                    self.label_le[n].setPixmap(QtGui.QPixmap(':/icons/cross.png'))
            else:
                 return
        # build visualization widget
        self.prFile = QtGui.QProgressBar(self.lineEdit[i])
        self.prFile.resize(self.lineEdit[i].width(), self.lineEdit[i].height())
        self.prFile.setMaximum(100)
        self.prFile.setTextVisible(True)
        self.prFile.setOrientation(QtCore.Qt.Horizontal)
        self.prFile.setTextDirection(QtGui.QProgressBar.TopToBottom)
        self.prFile.setFormat(self.tr('check file %p%'))
        layout = QtGui.QHBoxLayout(self.lineEdit[i])
        layout.addWidget(self.prFile,0,QtCore.Qt.AlignVCenter|QtCore.Qt.AlignLeft)
        layout.setSpacing(0)
        layout.setMargin(3)
        self.prFile.show()
        self.lineEdit[i].setText(filedb.replace('/', '\\'))
        for t in range(0, 101):
            time.sleep(0.005)
            self.prFile.setValue(t)
        time.sleep(0.5)
        self.prFile.setTextVisible(False)
        # check file reports
        if name_report[i] != name_str:
            msgError = self.tr('Required reports:<br>') + self.lfile[i].text()
            QtGui.QMessageBox.warning(self, self.tr('Incorrect file'), msgError)
            self.lineEdit[i].setText(None)
            self.label_le[i].setPixmap(QtGui.QPixmap(':/icons/cross.png'))
        else:
            self.label_le[i].setPixmap(QtGui.QPixmap(':/icons/tick.png'))
        f.close()
        if len(self.lineEdit[i].text()) > 30:
            self.lineEdit[i].setText(self.lineEdit[i].text() + '     ')
        self.file_changed()
        self.prFile.close()
        del self.prFile

    # ===================================
    def set_filter(self, i):
        data = self.filter_comboBox.itemData(i)
        if data == 'All':
            window.proxy.setFilterRegExp('')
            window.proxy.setFilterKeyColumn(1)
            window.tv.verticalHeader().show()
        else:
            window.proxy.setFilterRegExp(data)
            window.proxy.setFilterKeyColumn(1)
            window.tv.verticalHeader().hide()

# ===================================
if __name__ == "__main__":
    translator = QtCore.QTranslator()
    user_name = getpass.getuser()
    app = QtGui.QApplication(sys.argv)
    window = setup_win()
    take_option('load', 1)
    window.setInterface()
    load_splash()
    window.show()
    sys.exit(app.exec_())