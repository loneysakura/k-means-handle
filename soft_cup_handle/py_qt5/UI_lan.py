from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1129, 804)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(10, 10, 1141, 771))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 10, 511, 721))
        self.groupBox_3.setObjectName("groupBox_3")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.groupBox_3)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.listWidget = QtWidgets.QListWidget(self.groupBox_3)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout_2.addWidget(self.listWidget)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab)
        self.groupBox_4.setGeometry(QtCore.QRect(530, 360, 581, 371))
        self.groupBox_4.setObjectName("groupBox_4")
        self.textBrowserlog = QtWidgets.QTextBrowser(self.groupBox_4)
        self.textBrowserlog.setGeometry(QtCore.QRect(10, 30, 561, 291))
        self.textBrowserlog.setObjectName("textBrowserlog")
        self.progressBar = QtWidgets.QProgressBar(self.groupBox_4)
        self.progressBar.setGeometry(QtCore.QRect(90, 330, 491, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.lcdNumber = QtWidgets.QLCDNumber(self.groupBox_4)
        self.lcdNumber.setGeometry(QtCore.QRect(20, 330, 64, 23))
        self.lcdNumber.setObjectName("lcdNumber")
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        self.groupBox.setGeometry(QtCore.QRect(528, 42, 581, 311))
        self.groupBox.setObjectName("groupBox")
        self.pushButtonbrowse = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonbrowse.setGeometry(QtCore.QRect(70, 60, 113, 51))
        self.pushButtonbrowse.setObjectName("pushButtonbrowse")
        self.pushButtonclear = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonclear.setGeometry(QtCore.QRect(180, 60, 113, 51))
        self.pushButtonclear.setObjectName("pushButtonclear")

        self.pushButtonfirst = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonfirst.setGeometry(QtCore.QRect(290, 60, 113, 51))
        self.pushButtonfirst.setObjectName("pushButtonfirst")
        self.pushButtonload = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonload.setGeometry(QtCore.QRect(400, 60, 113, 51))
        self.pushButtonload.setObjectName("pushButtonload")
        self.pushButtonsecond = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonsecond.setGeometry(QtCore.QRect(70, 140, 113, 51))
        self.pushButtonsecond.setObjectName("pushButtonsecond")
        self.pushButtonthird = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonthird.setGeometry(QtCore.QRect(180, 140, 113, 51))
        self.pushButtonthird.setObjectName("pushButtonthird")
        self.pushButtonforth = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonforth.setGeometry(QtCore.QRect(290, 140, 113, 51))
        self.pushButtonforth.setObjectName("pushButtonfifth")

        self.pushButtonfifth = QtWidgets.QPushButton(self.groupBox)
        self.pushButtonfifth.setGeometry(QtCore.QRect(400, 140, 113, 51))
        self.pushButtonfifth.setObjectName("pushButtonfifth")

        self.comboBoxfiletype = QtWidgets.QComboBox(self.groupBox)
        self.comboBoxfiletype.setGeometry(QtCore.QRect(80, 220, 101, 26))
        self.comboBoxfiletype.setObjectName("comboBoxfiletype")

        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_2.setGeometry(QtCore.QRect(0, 10, 1091, 121))
        self.groupBox_2.setObjectName("groupBox_2")
        self.pushButton_confirm_idx = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_confirm_idx.setGeometry(QtCore.QRect(910, 80, 113, 32))
        self.pushButton_confirm_idx.setObjectName("pushButton_confirm_idx")
        self.pushButton_clear_idx = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_clear_idx.setGeometry(QtCore.QRect(910, 30, 113, 31))
        self.pushButton_clear_idx.setObjectName("pushButton_clear_idx")
        self.comboBox_wb = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_wb.setGeometry(QtCore.QRect(110, 30, 481, 26))
        self.comboBox_wb.setObjectName("comboBox_wb")
        self.comboBox_ws = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_ws.setGeometry(QtCore.QRect(110, 80, 481, 26))
        self.comboBox_ws.setObjectName("comboBox_ws")
        self.label_keyidx = QtWidgets.QLabel(self.groupBox_2)
        self.label_keyidx.setGeometry(QtCore.QRect(610, 30, 91, 31))
        self.label_keyidx.setObjectName("label_keyidx")
        self.label_range = QtWidgets.QLabel(self.groupBox_2)
        self.label_range.setGeometry(QtCore.QRect(610, 80, 91, 31))
        self.label_range.setObjectName("label_range")
        self.comboBox_x = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_x.setGeometry(QtCore.QRect(730, 30, 71, 26))
        self.comboBox_x.setObjectName("comboBox_x")
        self.comboBox_y = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_y.setGeometry(QtCore.QRect(800, 30, 91, 26))
        self.comboBox_y.setObjectName("comboBox_y")
        self.comboBox_r1 = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_r1.setGeometry(QtCore.QRect(730, 80, 71, 26))
        self.comboBox_r1.setObjectName("comboBox_r1")
        self.comboBox_r2 = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox_r2.setGeometry(QtCore.QRect(800, 80, 91, 26))
        self.comboBox_r2.setObjectName("comboBox_r2")
        self.checkBox_book = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox_book.setGeometry(QtCore.QRect(10, 30, 81, 31))
        self.checkBox_book.setChecked(False)
        self.checkBox_book.setObjectName("checkBox_book")
        self.checkBox_sheet = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox_sheet.setGeometry(QtCore.QRect(10, 80, 81, 31))
        self.checkBox_sheet.setChecked(False)
        self.checkBox_sheet.setObjectName("checkBox_sheet")
        self.groupBox_5 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_5.setGeometry(QtCore.QRect(0, 139, 1111, 591))
        self.groupBox_5.setObjectName("groupBox_5")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBox_5)
        self.tableWidget.setGeometry(QtCore.QRect(0, 30, 1111, 581))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.tabWidget.addTab(self.tab_2, "")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1129, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "电力客户行为分析小程序"))
        self.groupBox_3.setTitle(_translate("MainWindow", "excel路径表"))
        self.groupBox_4.setTitle(_translate("MainWindow", "显示区"))
        self.groupBox.setTitle(_translate("MainWindow", "功能区"))
        self.pushButtonbrowse.setText(_translate("MainWindow", "选择文件"))
        self.pushButtonclear.setText(_translate("MainWindow", "清除所有文件"))
        self.pushButtonfirst.setText(_translate("MainWindow", "功能一"))
        self.pushButtonload.setText(_translate("MainWindow", "加载excel文件"))
        self.pushButtonsecond.setText(_translate("MainWindow", "功能二"))
        self.pushButtonthird.setText(_translate("MainWindow", "功能三"))
        self.pushButtonforth.setText(_translate("MainWindow", "功能四"))
        self.pushButtonfifth.setText(_translate("MainWindow", "功能五"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Function_1"))
        self.groupBox_2.setTitle(_translate("MainWindow", "参数配置"))
        self.pushButton_confirm_idx.setText(_translate("MainWindow", "confirm"))
        self.pushButton_clear_idx.setText(_translate("MainWindow", "clear"))
        self.label_keyidx.setText(_translate("MainWindow", "  idx"))
        self.label_range.setText(_translate("MainWindow", " range"))
        self.checkBox_book.setText(_translate("MainWindow", "books"))
        self.checkBox_sheet.setText(_translate("MainWindow", "sheets"))
        self.groupBox_5.setTitle(_translate("MainWindow", "excel显示"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Function_2"))
