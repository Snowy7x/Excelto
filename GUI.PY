import sys
import PyQt5
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import QtCore, QtWidgets
import sqlite3
import pandas as pd
import os
import re
import logging

logging.getLogger(PyQt5.__name__).setLevel(logging.CRITICAL)


file_path = ""
col_name_list = []
main_types = []
branch_types = []
branch_types_m = []
rows_values = []

#clickedType = StringVar()
#writtedDiffer = StringVar()
searchTybes = []
o_searchTybes = []
search_text_list = None
search_text_list_2 = None
name_list = []
currSearchIndex = 0
currSearchValue = ""
currSearchBranchValue = ""
final_result = []


def is_number_regex(s):
    """ Returns True is string is a number. """
    if re.match("^\d+?\.\d+?$", s) is None:
        return s.isdigit()
    return True


app = QtWidgets.QApplication(sys.argv)
con = sqlite3.connect("tests.db")
cur = con.cursor()


######||>> GUI CODE <<||#######


class SelectWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Hello, world!'
        self.left = 10
        self.top = 10
        screen = self.screen()
        size = screen.size()
        self.width = 700
        self.height = 400
        self.initUI()
        self.setFixedSize(self.width, self.height)

    def initUI(self):
        self.setWindowTitle('Excel Reader & Conventor')
        self.setGeometry(self.left, self.top, self.width, self.height)
        style = "QPushButton"\
                "{"\
                "background-color : #CB0000;"\
                "font-size: 30px;"\
                "border: 4px solid black;"\
                "border-radius: 12px;"\
                "color: #D0D0D0;"\
                "}"\
                "QPushButton::hover"\
                "{"\
                "color: white;"\
                "font-size: 32px;"\
                "background-color : #E50000;"\
                "}"
        # Add BUTTON
        a_button = QPushButton('Add', self)
        a_button.setToolTip('Button')
        a_button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        a_button.setStyleSheet(style)
        a_button.clicked.connect(lambda: updateSearchType(self))
        a_button.move(int(self.width/3), int(self.height/1.3))
        a_button.resize(int(self.width/3.5), int(self.height/10))
        # DROPDOWN WIDGET
        self.cb = QComboBox(self)
        createSelectDropDown(self.cb)
        self.cb.currentIndexChanged.connect(self.selectionchange)
        self.cb.move(int(self.width/12), int(self.height/15))
        self.cb.resize(int(self.width/2.5), int(self.height/15))
        changeCurrValue(self.cb.currentText())
        # TEXT ENTRY WIDGET
        txt = QLineEdit(self)
        txt.textChanged.connect(self.text_changed)
        txt.setText("None")
        txt.move(int(self.width/1.8), int(self.height/15))
        txt.resize(int(self.width/2.5), int(self.height/15))

    def closeEvent(self, event):
        self.hide()
        event.accept()

    def selectionchange(self, i):
        changeCurrValue(self.cb.currentText())

    def text_changed(self, text):
        setCurrBranchSValue(text)


class MainWindow(QtWidgets.QWidget):

    global search_text_list
    global name_list
    switch_window = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.title = 'Excel Reader & Conventor'
        self.left = 10
        self.top = 10
        screen = self.screen()
        size = screen.size()
        self.width = size.width()
        self.height = size.height()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Reader & Conventor')
        self.setGeometry(self.left, self.top, self.width, self.height)
        style = "QPushButton"\
                "{"\
                "background-color : #CB0000;"\
                "font-size: 30px;"\
                "border: 4px solid black;"\
                "border-radius: 12px;"\
                "color: #D0D0D0;"\
                "}"\
                "QPushButton::hover"\
                "{"\
                "color: white;"\
                "font-size: 32px;"\
                "background-color : #E50000;"\
                "}"
        # SEARCH BUTTON
        s_button = QPushButton('Search', self)
        s_button.setToolTip('Button')
        s_button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        s_button.setStyleSheet(style)
        s_button.clicked.connect(searchResults)
        s_button.move(int(self.width/1.6), int(self.height/1.3))
        s_button.resize(int(self.width/3.5), int(self.height/20))
        # CLEAR BUTTON
        c_button = QPushButton('Clear', self)
        c_button.setToolTip('Button')
        c_button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        c_button.setStyleSheet(style)
        c_button.clicked.connect(clearSearch)
        c_button.move(int(self.width/7.6), int(self.height/1.3))
        c_button.resize(int(self.width/3.5), int(self.height/20))
        # DROPDOWN WIDGET
        self.cb = QComboBox(self)
        createDropDown(self.cb)
        self.cb.currentIndexChanged.connect(self.selectionchange)
        self.cb.move(int(self.width/10), int(self.height/15))
        self.cb.resize(int(self.width/3.5), int(self.height/20))
        # TEXT SEARCH
        self.textbox = QLabel('Search input will be here', self)
        self.textbox.setFont(QFont('Arial', 14))
        self.textbox.move(int(self.width/22), int(self.height/6))
        self.textbox.resize(int(self.width/1.1), int(self.height/2))
        setSearchText(self.textbox)
        comboBoxCallBack(self.cb, False)

    def selectionchange(self, i):
        comboBoxCallBack(self.cb)
        print("Current index", i, "selection changed ", self.cb.currentText())


class WindowTwo(QtWidgets.QWidget):

    global search_text_list
    global name_list
    switch_window = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.title = 'Excel Reader & Conventor'
        self.left = 10
        self.top = 10
        screen = self.screen()
        size = screen.size()
        self.width = size.width()/4
        self.height = size.height()/4
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Reader & Conventor')
        self.setGeometry(self.left, self.top, self.width, self.height)
        style = "QPushButton"\
                "{"\
                "background-color : #CB0000;"\
                "font-size: 20px;"\
                "border: 4px solid black;"\
                "border-radius: 12px;"\
                "color: #D0D0D0;"\
                "}"\
                "QPushButton::hover"\
                "{"\
                "color: white;"\
                "font-size: 22px;"\
                "background-color : #E50000;"\
                "}"
        # SEARCH BUTTON
        s_button = QPushButton('Search', self)
        s_button.setToolTip('Button')
        s_button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        s_button.setStyleSheet(style)
        s_button.clicked.connect(lambda: searchResults(True))
        s_button.move(int(self.width/1.6), int(self.height/1.3))
        s_button.resize(int(self.width/3.5), int(self.height/8))
        # CLEAR BUTTON
        c_button = QPushButton('Clear', self)
        c_button.setToolTip('Button')
        c_button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        c_button.setStyleSheet(style)
        c_button.clicked.connect(lambda: clearSearch(False, None, True))
        c_button.move(int(self.width/7.6), int(self.height/1.3))
        c_button.resize(int(self.width/3.5), int(self.height/8))
        # DROPDOWN WIDGET
        self.cb = QComboBox(self)
        createDropDown(self.cb)
        self.cb.currentIndexChanged.connect(self.selectionchange)
        self.cb.move(int(self.width/10), int(self.height/15))
        self.cb.resize(int(self.width/3.5), int(self.height/8))
        # TEXT SEARCH
        self.textbox = QLabel('Search input will be here', self)
        self.textbox.setFont(QFont('Arial', 8))
        self.textbox.move(int(self.width/22), int(self.height/6))
        self.textbox.resize(int(self.width/1.1), int(self.height/2))
        setSearchText2(self.textbox)
        comboBoxCallBack(self.cb, False)

    def selectionchange(self, i):
        comboBoxCallBack(self.cb)
        print("Current index", i, "selection changed ", self.cb.currentText())


class Main(QtWidgets.QWidget):

    global file_path
    switch_window = QtCore.pyqtSignal()

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.setWindowTitle('Excel Reader & Conventor')

        layout = QtWidgets.QGridLayout()
        screen = self.screen()
        size = screen.size()
        self.width = size.width()
        self.height = size.height()
        self.button = QtWidgets.QPushButton('Choose Excel file')
        self.button.clicked.connect(self.chooseFile)
        self.button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        self.button.setStyleSheet("QPushButton"
                                  "{"
                                  "background-color : #CB0000;"
                                  "font-size: 30px;"
                                  "border: 4px solid black;"
                                  "color: #D0D0D0;"
                                  "}"
                                  "QPushButton::hover"
                                  "{"
                                  "color: white;"
                                  "font-size: 34px;"
                                  "background-color : #E50000;"
                                  "}")
        self.button.setMaximumHeight(100)
        self.button.setMaximumWidth(600)

        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.button.setSizePolicy(sizePolicy)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def chooseFile(self):
        fname = QFileDialog.getOpenFileName(self, 'Choose Excel file')
        if fname[0]:
            file_path = fname[0]
            prepareDB(fname[0], self)


class ResultTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Excel Reader & Conventor')
        ##################################
        searchAct = QAction(QIcon('gif-search.gif'), 'Search', self)
        searchAct.setShortcut('Ctrl+F')
        searchAct.triggered.connect(getController().show_window_two)

        self.toolbar = QToolBar('Search')
        self.toolbar.addAction(searchAct)
        #################################
        screen = self.screen()
        size = screen.size()
        self.width = size.width()
        self.height = size.height()-100

        types = getMainTypes()[2:]
        self.layout = QVBoxLayout(self)
        self.layout.setMenuBar(self.toolbar)

        self.names = QTabWidget()
        n_tab = QWidget()
        n_tab.layout = QVBoxLayout(self)
        table = QTableWidget()
        branchs = getBranchs()
        for branch in branchs:
            if branch == "None":
                branchs.remove(branch)
                try:
                    del branchs[branchs.index(branch)]
                except Exception as e:
                    pass
        rows = getFinalResult()
        cols = getTypeBranchs(getMainTypes()[1])

        table.setRowCount(len(rows)+1)
        # Column count
        table.setColumnCount(len(cols))
        for x, row in enumerate(rows):
            for y, col in enumerate(cols):
                table.setItem(0, y, QTableWidgetItem(cols[y]))
        for x, row in enumerate(rows):
            for y, value in enumerate(row):
                if y >= len(cols)+2:
                    break
                branch = branchs[y]
                if branch in cols:
                    table.setItem(x+1, y-2, QTableWidgetItem(row[y]))

        n_tab.layout.addWidget(table)
        n_tab.setLayout(n_tab.layout)
        self.names.addTab(n_tab, getMainTypes()[1])

        self.names.setMaximumWidth(int(self.width/6))
        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tabs.setMaximumWidth(int((self.width/6)*5))
        self.tabs.resize(300, 200)
        self.allTabs = []

        # Add tabs
        for index, value in enumerate(types):
            tab = QWidget()
            tab.layout = QVBoxLayout(self)
            table = QTableWidget()

            rows = getFinalResult()
            cols = getTypeBranchs(value)

            table.setRowCount(len(rows)+1)
            # Column count
            table.setColumnCount(len(cols))
            print(f'{len(branchs)}   ===   {len(rows[0])}')
            for x, row in enumerate(rows):
                for y, col in enumerate(cols):
                    table.setItem(0, y, QTableWidgetItem(col))
            for x, row in enumerate(rows):
                for y, branch in enumerate(branchs):
                    if branch in cols and branch != 'None':
                        col_i = cols.index(branch)
                        print(f'{y} +> {branchs[y]}')
                        table.setItem(x+1, col_i, QTableWidgetItem(row[y]))
            tab.layout.addWidget(table)
            tab.setLayout(tab.layout)
            self.allTabs.append(tab)
            self.tabs.addTab(tab, value)

        # Add tabs to widget
        self.layout.setDirection(QBoxLayout.LeftToRight)
        self.layout.addWidget(self.names)
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
        self.resize(300, 300)


class ResultTab2(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Filtered Results')
        ###################################
        searchAct = QAction(QIcon('gif-search.gif'), 'Search', self)
        searchAct.setShortcut('Ctrl+F')
        searchAct.triggered.connect(getController().show_window_two)
        self.toolbar = QToolBar('Search')
        self.toolbar.addAction(searchAct)
        ##################################
        screen = self.screen()
        size = screen.size()
        self.width = size.width()
        self.height = size.height()-100

        types = getMainTypes()[2:]
        self.layout = QVBoxLayout(self)
        self.layout.setMenuBar(self.toolbar)

        self.names = QTabWidget()
        n_tab = QWidget()
        n_tab.layout = QVBoxLayout(self)
        table = QTableWidget()
        branchs = getBranchs()
        for branch in branchs:
            if branch == "None":
                branchs.remove(branch)
                try:
                    del branchs[branchs.index(branch)]
                except Exception as e:
                    pass
        rows = getFinalResult()
        cols = getTypeBranchs(getMainTypes()[1])

        table.setRowCount(len(rows)+1)
        # Column count
        table.setColumnCount(len(cols))
        for x, row in enumerate(rows):
            for y, col in enumerate(cols):
                table.setItem(0, y, QTableWidgetItem(cols[y]))
        for x, row in enumerate(rows):
            for y, value in enumerate(row):
                if y >= len(cols)+2:
                    break
                branch = branchs[y]
                if branch in cols:
                    table.setItem(x+1, y-2, QTableWidgetItem(row[y]))

        n_tab.layout.addWidget(table)
        n_tab.setLayout(n_tab.layout)
        self.names.addTab(n_tab, getMainTypes()[1])

        self.names.setMaximumWidth(int(self.width/6))
        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tabs.setMaximumWidth(int((self.width/6)*5))
        self.tabs.resize(300, 200)
        self.allTabs = []

        # Add tabs
        for index, value in enumerate(types):
            tab = QWidget()
            tab.layout = QVBoxLayout(self)
            table = QTableWidget()

            rows = getFinalResult()
            cols = getTypeBranchs(value)

            table.setRowCount(len(rows)+1)
            # Column count
            table.setColumnCount(len(cols))
            for x, row in enumerate(rows):
                for y, col in enumerate(cols):
                    table.setItem(0, y, QTableWidgetItem(col))
            for x, row in enumerate(rows):
                for y, branch in enumerate(branchs):
                    if branch in cols and branch != 'None':
                        col_i = cols.index(branch)
                        table.setItem(x+1, col_i, QTableWidgetItem(row[y]))
            tab.layout.addWidget(table)
            tab.setLayout(tab.layout)
            self.allTabs.append(tab)
            self.tabs.addTab(tab, value)

        # Add tabs to widget
        self.layout.setDirection(QBoxLayout.LeftToRight)
        self.layout.addWidget(self.names)
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
        self.resize(self.width/4, self.height/4)


class Controller:

    def __init__(self):
        self.select_on = False
        pass

    def show_main(self):
        self.main = Main()
        screen = self.main.screen()
        size = screen.size()
        self._width = size.width()
        self._height = size.height()-100
        self.main.switch_window.connect(self.show_select)
        self.main.setFixedSize(self._width, self._height)
        self.main.showMaximized()

    def show_select(self):
        self.show_select = MainWindow()
        screen = self.show_select.screen()
        size = screen.size()
        self._width = size.width()
        self._height = size.height()-100
        self.show_select.switch_window.connect(self.show_window_two)
        self.main.close()
        self.show_select.setFixedSize(self._width, self._height)
        self.show_select.showMaximized()

    def show_select_br(self):
        global select_on
        self.select_br = SelectWindow()
        self.select_on = True
        self.select_br.show()

    def show_window_two(self):
        saveSearch()
        global o_searchTybes
        global searchTybes
        o_searchTybes = searchTybes
        self.window_two = WindowTwo()
        self.window_two.show()

    def show_Results(self):
        self.results = ResultTab()
        screen = self.results.screen()
        size = screen.size()
        self._width = size.width()
        self._height = size.height()-100
        # self.show_select.close()
        self.results.setFixedSize(self._width, self._height)
        self.results.showMaximized()
        pass

    def show_Results2(self):
        self.results2 = ResultTab2()
        self.results2.show()
        pass


controller = Controller()


def getController():
    global controller
    return controller

#######||>> DATABASE CODE <<||######


def saveSearch():
    global searchTybes
    global o_searchTybes
    slist = searchTybes
    o_searchTybes = slist


def showError(title="Error", error="Something wrong happened!", details='No'):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)

    msg.setText(title)
    msg.setInformativeText(error)
    msg.setWindowTitle(title)
    if details != "No":
        msg.setDetailedText(details)
    msg.setStandardButtons(QMessageBox.Ok)
    retval = msg.exec_()
    print("value of pressed message box button:", retval)


def getFinalResult():
    global final_result
    return final_result


def getTypeBranchs(typ):
    global main_types
    global branch_types
    index = main_types.index(typ)
    return branch_types[index]


def getResultByTybe(typ):
    global final_result
    global main_types
    global branch_types
    global branch_types_m
    index = main_types.index(typ)
    branchs = branch_types[index]
    typ_results = [["None"] * len(branchs)] * len(final_result)
    for x, row in enumerate(final_result):
        for y, value in enumerate(row):
            branch = branch_types_m[y]
            if branch in branchs:
                typ_results[x][branchs.index(branch)] = value
        typ_results[x] = tuple(typ_results[x])
    #showError("Error", "columns in results does not match default")


def clearSearch(closeWin=False, win=None, search_2=False):
    global searchTybes
    global o_searchTybes
    global search_text_list
    global search_text_list_2
    if search_text_list != None:
        search_text_list.setText("")
    if search_2:
        search_text_list.setText('')
        searchTybes = o_searchTybes
    else:
        searchTybes = []
    if closeWin and win != None:
        win.close()
    pass


def searchResults(is2=False):
    global searchTybes
    global final_result
    if len(searchTybes) <= 0:
        showError("Not Index!", "Please add some search indexs")
        #messagebox.showerror("Not Index!", "Please add some search indexs")
        return

    search_cmd = 'SELECT * FROM MAIN WHERE'
    for s in searchTybes:
        try:
            value = s.split(':')[0]
            s_value = s.split(':')[1]
        except Exception as e:
            print(e)
            continue
        if search_cmd != 'SELECT * FROM MAIN WHERE':
            search_cmd += " AND"

        if s_value == 'None' or s_value == '' or s_value is None:
            search_cmd += " [{0}] NOT LIKE 'None%'".format(value)
            search_cmd += " AND [{0}] IS NOT NULL".format(value)
        else:
            search_cmd += " [{0}] LIKE '{1}'".format(value, s_value)

    print(search_cmd)

    cur.execute(search_cmd)
    # cur.execute('SELECT * FROM MAIN WHERE "{0}" LIKE "{1}"'.format(s_type, name))

    rows = cur.fetchall()
    if len(rows) > 0:
        print("search is done and found some results")
        final_result = rows
        getResultByTybe(main_types[1])
        # clearResults(rows)
        if is2:
            controller.show_Results2()
        else:
            controller.show_Results()
    else:
        showError(
            "Not Found!", "There is nothing match your search please check your inputs and try again")


def setCurrBranchSValue(txt):
    global currSearchBranchValue
    currSearchBranchValue = txt


def changeCurrValue(value):
    global currSearchValue
    currSearchValue = value


def createNewDB(widget):
    global con
    global cur
    print("Creating new DataBase")
    con = sqlite3.connect("main.db")
    cur = con.cursor()
    try:
        cur.execute('''DROP TABLE MAIN''')
        cur.commit()
    except Exception as e:
        pass

    create_cmd = '''CREATE TABLE MAIN('''
    insert_cmd = '''INSERT INTO MAIN('''
    global rows_values

    # creating new table
    x = 0
    valid_x = []
    cols = []
    for col in branch_types_m:
        if col == "None":
            for row in rows_values:
                valid_x.append(x)
            x += 1
            continue

        # setting up the indexs
        index = 0
        m_index = 0
        for b in branch_types:
            if col in b:
                m_index = branch_types.index(b)
                index = b.index(col)
        _temp = col
        if col.lower() in cols:
            while col.lower() in cols:
                for b in branch_types:
                    if col in b and col.lower() in cols:
                        col = col + "-" + main_types[m_index]
                j = 0
                o_col = col
                while col.lower() in cols:
                    if col.lower() in cols:
                        col = f"{o_col}-({j})"
                    j += 1

        branch_types[m_index][index] = col
        branch_types_m[branch_types_m.index(_temp)] = col

        cols.append(col.lower())
        if create_cmd == '''CREATE TABLE MAIN(''':
            create_cmd += f'''[{col}] text'''
            insert_cmd += f'''[{col}] '''
        else:
            create_cmd += f''', [{col}] text'''
            insert_cmd += f''', [{col}] '''
        x += 1

    create_cmd += ''')'''
    insert_cmd += ''') VALUES('''
    cur.execute(create_cmd)
    con.commit()
    # inserting the values
    for row in rows_values:
        y = 0
        insert = insert_cmd
        x = 0
        for value in row:
            if y in valid_x:
                pass
            else:
                if value == None or value == "None":
                    value = "None"
                if x == 0:
                    insert += f''''{value.replace("'", "")}' '''
                else:
                    insert += f''','{value.replace("'", "")}' '''
                x += 1
            y += 1
        insert += ''')'''
        try:
            cur.execute(insert)
            con.commit()
        except Exception as e:
            print("ERROR: " + str(e))
            print(str(insert))
    widget.switch_window.emit()
    pass


def clearData(widget):
    global col_name_list
    global main_types
    global branch_types
    global branch_types_m
    global rows_values
    print('Clearing the data')
    i = 0
    x = 0
    y = 0
    cuts = []
    unnamed_list = []
    _list = []
    new_list = []
    for col in col_name_list:
        if "Unnamed" in col:
            unnamed_list.append(col)
            i += 1
        else:
            cuts.append(y)
            col = str(col).replace('\n', '')
            new_list.append(col)
            x += 1
        y += 1

    main_types = new_list

    cur.execute('SELECT * FROM Sheet1')
    rows = cur.fetchall()
    y = 0
    brachs = []

    x = 0
    for row in rows[0]:
        row = f"{row}"
        if y in cuts and cuts[0] != y:
            _list.append(brachs)
            brachs = []
        if row == 'None' or row is None:

            for z in range(1, 4):
                if row == 'None' or row is None:
                    row = rows[z][y]

        if row == 'None' or row is None:
            pass  # row = "--"

        row = str(row).replace(".", '')
        row = str(row).replace('\n', '')

        brachs.append(row)
        branch_types_m.append(row)

        y += 1
    _list.append(brachs)
    branch_types = _list

    # rows values cleaner
    rows_values = rows
    for j in range(0, 4):
        rows_values.remove(rows_values[0])
    print(
        f"Unnamed cols: {i},  Named cols: {x}, cols count: {len(col_name_list)},  new cols count {len(new_list)}")
    createNewDB(widget)


def prepareDB(path, widget):
    global main_types
    global name_list
    global file_path
    file_path = path
    global cur
    print("Preparing The Database")
    currdir = os.getcwd()
    if len(file_path) > 0:
        print("You chose: %s" % file_path)

    wb = pd.ExcelFile(file_path)
    for sheet in wb.sheet_names:
        df = pd.read_excel(file_path)
        df.to_sql(sheet, con, index=False, if_exists="replace")

    global col_name_list

    cur.execute("SELECT * FROM Sheet1")
    col_name_list = [tuple[0] for tuple in cur.description]

    clearData(widget)
    name_list = main_types[1:]
    # CreateDropDown(name_list=main_types[1:])


def createDropDown(cb):
    cb.addItems(main_types[1:])
    cb.setCurrentText("Choose a type")
    pass


def createSelectDropDown(cb: QComboBox):
    cb.addItems(branch_types[currSearchIndex])
    pass


def setSearchText(s_text):
    global search_text_list
    search_text_list = s_text
    pass


def setSearchText2(s_text):
    global search_text_list_2
    search_text_list_2 = s_text
    pass


def updateSearchType(widget):
    global currSearchValue
    global search_text_list
    global currSearchBranchValue
    searchTybes[len(
        searchTybes) - 1] = (currSearchValue + f":{currSearchBranchValue}")
    search_text_list.setText("")
    if search_text_list_2 != None:
        search_text_list_2.setText('')
    for s in searchTybes:
        if search_text_list_2 != None:
            search_text_list_2.setText(search_text_list_2.text()+f"\n{s}")
        search_text_list.setText(search_text_list.text()+f"\n{s}")
    widget.close()


def comboBoxCallBack(cb: QComboBox, open=True):
    global currSearchIndex
    global current_search
    global searchTybes
    value = cb.currentText()
    index = cb.currentIndex() + 1
    if len(searchTybes) > 0:
        for s in searchTybes:
            if ':' not in s:
                print("Found Wrong Format")
                searchTybes.remove(s)
                try:
                    del searchTybes[s]
                except Exception:
                    pass
                continue
            v = s.split(':')[0]
            for b in branch_types[index]:
                if v in b:
                    print("Found Dublicate")
                    searchTybes.remove(s)
                    try:
                        del searchTybes[s]
                    except Exception:
                        pass

    v = branch_types[index][0]
    searchTybes.append(v)
    currSearchIndex = index
    if open:
        controller.show_select_br()


def getBranchs():
    return branch_types_m


def getBranchTypes():
    return branch_types


def getMainTypes():
    return main_types


def main():
    controller.show_main()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
