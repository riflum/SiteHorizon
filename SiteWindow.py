import sys
from itertools import chain
import logging

from docx import Document
from docx.shared import Cm, Pt, Inches

from PyQt5.QtWidgets import  (QFileDialog, QFrame, QGridLayout, QGroupBox, QComboBox, QLineEdit, QWidget, 
                                QPushButton, QVBoxLayout, QDialog, QMessageBox, QTableView, QCompleter, 
                                QMainWindow, QLabel, QApplication, QTextEdit, QDataWidgetMapper)
from PyQt5.QtCore import QRect, QModelIndex, QObject, pyqtSignal, Qt, QEvent, QSortFilterProxyModel, QDate, QDateTime
from PyQt5.QtSql import QSqlQuery, QSqlTableModel, QSqlRecord, QSqlQueryModel
from PyQt5 import uic, QtCore

import connection
import sqlite3

import ArrayList as list

uifile = "UI/SiteWindow.ui"
# # Add Blok Petak
# uifile = "UI/SiteWindow_Blok_Petak.ui"
Ui_MainWindow, QtBaseClass = uic.loadUiType(uifile)

def parseRange(rng):
    parts = rng.split("-")
    if 1 > len(parts) > 2:
        raise ValueError ("Bad range: '%s'" %s (rng, ))
    parts = [int (i) for i in parts]
    start = parts[0]
    end = start if len(parts) == 1 else parts[1]
    if start > end:
        end, start = start, end
    return range(start, end + 1)

def parseRangeList(rngs):
    return sorted(set(chain(*[parseRange(rng) for rng in rngs.split(',')])))

def proxyModelRowIndex(model, text, textLanguage=1):
    proxy = QSortFilterProxyModel()
    proxy.setSourceModel(model)
    proxy.setFilterFixedString(text)
    matchingIndex = proxy.mapToSource(proxy.index(0,0))
    rowKindObs = matchingIndex.row()
    result = model.data(model.index(rowKindObs, textLanguage))
    return result
    

def initializeModel(model, table, filter):
    model.setTable(table)

    model.setEditStrategy(QSqlTableModel.OnManualSubmit)
    model.setFilter(filter)
    model.select()

def initializeTableEditModel(model, table):
    model.setTable(table)
    model.setEditStrategy(QSqlTableModel.OnFieldChange)
    model.select()

class FocusOutFilter(QObject):
    focusOut = pyqtSignal()

    def eventFilter(self, widget, event):
        if event.type() == QEvent.FocusOut:
            # print("--eventFilter() focus_out on: " + widget.objectName())
            logging.info("--eventFilter() focus_out on: %s" + widget.objectName())
            self.focusOut.emit()
            # return false so there are not two or more blinking cursor
            # focus will not out
            return False
        else:
            return False

class FocusInFilter(QObject):
    focusIn = pyqtSignal()

    def eventFilter(self, widget, event):
        if event.type() == QEvent.FocusIn:
            # print("--eventFilter() focus_out on: " + widget.objectName())
            logging.info("--eventFilter() focus_out on: %s" + widget.objectName())
            self.focusIn.emit()
            # return false so there are not two or more blinking cursor
            # focus will not out
            return False
        else:
            return False

def tableViewCombobox(tableView, combobox, width2):
    tableView = QTableView()
    tableView.verticalHeader().hide()
    tableView.horizontalHeader().hide()
    combobox.setView(tableView)
    tableView.hideColumn(1)
    tableView.setColumnWidth(0, 50)
    tableView.setColumnWidth(2, width2)

def populate(model, table, combobox, completer):
    """ populate the model in the combobox with completer """

    initializeModel(model, table, '')
    combobox.setModel(model)
    combobox.setModelColumn(1)
    combobox.setCurrentIndex(-1)
    completer.setModel(model)
    completer.setCompletionColumn(1)
    combobox.setCompleter(completer)

def populatePlusTable(model, table, combobox, completer, tableView, width2):
    """ populate the model in the combobox with tableview for combobox and completer """
    # textureClassModel = QSqlTableModel()
    # 1st initialize the model
    initializeModel(model, table, '')
    # set the combobox model
    combobox.setModel(model)
    # set the column index to show in the combobox
    combobox.setModelColumn(0)
    # set the default value to empty/blank
    combobox.setCurrentIndex(-1)
    # set the completer model
    completer.setModel(model)
    # set the combobox completer
    combobox.setCompleter(completer)
    # sort the combobox from '0' column
    combobox.model().sort(0, Qt.SortOrder(Qt.AscendingOrder))

    # set the completer table view
    completerTableView = QTableView()
    ## hidden the vertical and horizontal header of the tableview in completer
    completerTableView.verticalHeader().hide()
    completerTableView.horizontalHeader().hide()
    # set the table view with setPopup function
    completer.setPopup(completerTableView)
    # hide TEXT_INDO colum - index = 1; TEXT_ENGLISH - index = 2
    completerTableView.hideColumn(1)
    # set the 1st column width
    completerTableView.setColumnWidth(0, 40)
    # set the 2nd column width
    completerTableView.setColumnWidth(2, 120)
    # set the width of the completer to beyond the width of the combobox
    completer.popup().setMinimumWidth(165)
    completer.popup().setMinimumHeight(30)

    # set the tableview for dropdown combobox
    tableView = QTableView()
    ## hidden the vertical and horizontal header of the tableview in combobox drop-down
    tableView.verticalHeader().hide()
    tableView.horizontalHeader().hide()
    # set the table view with setPopup function
    combobox.setView(tableView)
    # hide TEXT_INDO colum - index = 1, TEXT_ENGLISH = index = 2 (tableView.hideColumn(index))
    tableView.hideColumn(1)
    # set the width of the column
    tableView.setColumnWidth(0, 50)
    # width2 = set the TEXT_INDO/ENGLISH width not the combobox width
    tableView.setColumnWidth(2, width2)

class SiteWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent = None):
        super().__init__()
        Ui_MainWindow.__init__(self)
        # setup the SiteWindow.ui
        self.setupUi(self)
        ### initialize the site model
        self.siteModel = QSqlTableModel()
        siteFilter = ''
        initializeModel(self.siteModel, "Site", siteFilter)

        ### initialise Edit Table Site Model
        self.siteModelTableEdit = QSqlTableModel()
        initializeTableEditModel(self.siteModelTableEdit, "Site")

        ## initialize the site record for inserting data
        self.siteRecord = QSqlRecord()
        self.siteRecord = self.siteModel.record()

        ### initialize completer, model, style and mapper
        self.initCompleter()
        self.populateData()
        self.setComboBoxStyle()
        self.initMapper()

        # initialize the events
        self.setEvents()

    def setEvents(self):
        """ populate the event """
        logging.info("set Site Windows Events")

        self.Reload_PB.clicked.connect(self.check)
        self.Provinsi_CB.currentTextChanged.connect(self.provinsiTextChanged)
        self.Kabupaten_CB.currentTextChanged.connect(self.kabupatenKotaTextChanged)
        self.Kecamatan_CB.currentTextChanged.connect(self.kecamatanTextChanged)

        self.Number_Form_CB.currentTextChanged.connect(self.numberFormTextChanged)
        self.Lower_Limit_1_LE.textChanged.connect(self.lowerLimit1TextChanged)
        # self.Lower_Limit_2_LE.textChanged.connect(self.lowerLimit2TextChanged)
        # self.Lower_Limit_3_LE.textChanged.connect(self.lowerLimit3TextChanged)
        # self.Lower_Limit_4_LE.textChanged.connect(self.lowerLimit4TextChanged)
        

        # Button
        self.Copy_PB.clicked.connect(self.copyData)
        self.Insert_PB.clicked.connect(self.insertData)
        self.Update_PB.clicked.connect(self.updateData)
        self.Delete_PB.clicked.connect(self.deleteData)
        self.Show_Table_PB.clicked.connect(self.showTable)

        self.First_PB.clicked.connect(self.toFirst)
        self.Previous_PB.clicked.connect(self.toPrevious)
        self.Next_PB.clicked.connect(self.toNext)
        self.Last_PB.clicked.connect(self.toLast)

        # self.Action_Export_to_Docx.triggered.connect(self.exportToDocx)

        # Focus Out
        self.greatGroupFocusOut = FocusOutFilter()
        self.Great_Group_CB.installEventFilter(self.greatGroupFocusOut)
        self.subGroupFocusOut = FocusOutFilter()
        self.Sub_Group_CB.installEventFilter(self.subGroupFocusOut)
        self.pscFocusOut = FocusOutFilter()
        self.PSC_CB.installEventFilter(self.pscFocusOut)
        self.mineralFocusOut = FocusOutFilter()
        self.Mineral_CB.installEventFilter(self.mineralFocusOut)
        self.temperatureFocusOut = FocusOutFilter()
        self.Temperature_CB.installEventFilter(self.temperatureFocusOut)
    
        self.greatGroupFocusOut.focusOut.connect(self.greatGroupFocusLost)
        self.subGroupFocusOut.focusOut.connect(self.subGroupFocusLost)
        self.pscFocusOut.focusOut.connect(self.pscFocusLost)
        self.mineralFocusOut.focusOut.connect(self.mineralFocusLost)
        # self.temperatureFocusOut.focusOut.connect(self.temperatureFocusLost)

    def exportToDocx(self):
        print("export")
         # make dialog for export the docx

        frameStyle = QFrame.Sunken | QFrame.Panel

        self.exportDialog = QDialog()
        self.gridLayout = QGridLayout()
        self.languageLabel = QLabel("Language:")
        self.languageComboBox = QComboBox()
        self.languageComboBox.addItem("")
        self.languageComboBox.addItem("INDONESIA")
        self.languageComboBox.addItem("ENGLISH")
        self.languageComboBox.setCurrentIndex(2)
        self.rangeLabel = QLabel("Range Number Form:")
        self.rangeLineEdit = QLineEdit()
        self.okPushButton = QPushButton("OK")
        self.cancelPushButton = QPushButton("Cancel")
        self.saveDocxFile = QLineEdit()
        self.saveDocxFile.setReadOnly(True)
        self.savePushButton = QPushButton("Docx Location")

        self.gridLayout.setColumnStretch(1, 1)
        self.gridLayout.setColumnMinimumWidth(1, 150)
        self.gridLayout.addWidget(self.languageLabel, 0, 0)
        self.gridLayout.addWidget(self.languageComboBox, 0, 1)
        self.gridLayout.addWidget(self.rangeLabel, 1, 0)
        self.gridLayout.addWidget(self.rangeLineEdit, 1, 1)
        self.gridLayout.addWidget(self.saveDocxFile, 2, 0, 2, 2)
        self.gridLayout.addWidget(self.savePushButton, 4, 0, 1, 2)
        self.gridLayout.addWidget(self.okPushButton, 5, 0)
        self.gridLayout.addWidget(self.cancelPushButton, 5, 1)
        self.exportDialog.setLayout(self.gridLayout)
        self.okPushButton.clicked.connect(self.okClicked)
        self.cancelPushButton.clicked.connect(self.exportDialog.close)
        self.savePushButton.clicked.connect(self.setSaveDocx)
        self.exportDialog.setWindowTitle("Export to Docx")
        self.exportDialog.exec()

    def copyData(self):
        """ Event to copy the entire fields from another number form selected
            so that cant be easily edited and inserted to database """

        # print("Copy")
        logging.debug("Copy the data")
        # get the current index of selected number form
        indexCopy = self.Copy_CB.currentIndex()

        logging.info("Copy form %s", indexCopy)

        ### set the current text/data from another number form
        self.SPT_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,2))))
        self.Provinsi_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,3))))
        self.Kabupaten_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,4))))
        self.Kecamatan_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,5))))
        self.Desa_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,6))))

        self.Date_CB.setDate(QDate.fromString(str(self.siteModel.data(self.siteModel.index(indexCopy,7))), "dd/MM/yyyy"))
        self.Initial_Name_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,8))))
        self.Observation_Number_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,9))))
        self.Kind_Observation_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,10))))
        self.UTM_Zone_1_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,11))))
        self.UTM_Zone_2_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,12))))
        self.X_East_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,13))))
        self.Y_North_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,14))))
        self.Elevation_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,15))))
        self.Landform_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,16))))
        self.Relief_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,17))))
        self.Slope_Percentage_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,18))))
        self.Slope_Shape_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,19))))
        self.Slope_Position_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,20))))
        self.Parent_Material_Accumulation_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,21))))
        self.Parent_Material_Lithology_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,22))))
        self.Parent_Material_Lithology_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,23))))
        self.Drainage_Class_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,24))))
        self.Permeability_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,25))))
        self.Runoff_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,26))))
        self.Erosion_Type_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,27))))
        self.Erosion_Degree_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,28))))
        self.Effective_Soil_Depth_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,29))))
        self.Land_Cover_Main_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,30))))
        self.Land_Cover_Climate_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,31))))
        self.Land_Cover_Vegetation_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,32))))
        self.Land_Cover_Performance_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,33))))
        self.Land_Cover_Main_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,34))))
        self.Land_Cover_Climate_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,35))))
        self.Land_Cover_Vegetation_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,36))))
        self.Land_Cover_Performance_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,37))))
        self.Soil_Moisture_Regime_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,38))))
        self.Soil_Temperature_Regime_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,39))))
        self.Epipedon_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,40))))
        self.Upper_Limit_1_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,41))))
        self.Lower_Limit_1_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,42))))
        self.Sub_Surface_Horizon_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,43))))
        self.Upper_Limit_2_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,44))))
        self.Lower_Limit_2_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,45))))
        self.Sub_Surface_Horizon_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,46))))
        self.Upper_Limit_3_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,47))))
        self.Lower_Limit_3_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,48))))
        self.Sub_Surface_Horizon_3_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,49))))
        self.Upper_Limit_4_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,50))))
        self.Lower_Limit_4_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,51))))
        self.Sub_Surface_Horizon_4_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,52))))
        self.Upper_Limit_5_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,53))))
        self.Lower_Limit_5_LE.setText(str(self.siteModel.data(self.siteModel.index(indexCopy,54))))
        self.Taxonomy_Year_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,55))))
        self.Great_Group_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,56))))
        self.Sub_Group_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,57))))
        self.PSC_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,58))))
        self.Mineral_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,59))))
        self.Reaction_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,60))))
        self.Temperature_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(indexCopy,61))))

    def deleteData(self):
        """ method to delete the row in the model and database table """

        print("delete")
        logging.info("Delete Database button clicked")

        # get the current text/number/value from number form combobox
        numberFormValue = self.Number_Form_CB.currentText()
        # find the row index of the desire number form in the model
        row = self.Number_Form_CB.findText(numberFormValue)

        # remove the row
        self.siteModel.removeRow(row)
        
        # submit the removeRow/delete function
        self.siteModel.submitAll()

        # Reset auto increment
        query = QSqlQuery()
        query.exec("UPDATE SQLITE_SEQUENCE SET SEQ=0 WHERE NAME='Site'")
        
    def updateData(self):
        """ Update the data from the fields to model and database tables """

        # print("update")
        logging.info("Update Database button clicked")

        # get the row index from the current value/number of number form
        row = self.Number_Form_CB.findText(self.Number_Form_CB.currentText())
        print("row:", row)

        ### update the data to the model and database 
        # self.siteModel.setData(self.siteModel.index(row, 1), self.Number_Form_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 2), self.SPT_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 3), self.Provinsi_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 4), self.Kabupaten_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 5), self.Kecamatan_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 6), self.Desa_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 7), str(self.Date_CB.date().toString("dd/MM/yyyy")))
        self.siteModel.setData(self.siteModel.index(row, 8), self.Initial_Name_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 9), self.Observation_Number_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 10), self.Kind_Observation_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 11), self.UTM_Zone_1_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 12), self.UTM_Zone_2_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 13), self.X_East_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 14), self.Y_North_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 15), self.Elevation_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 16), self.Landform_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 17), self.Relief_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 18), self.Slope_Percentage_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 19), self.Slope_Shape_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 20), self.Slope_Position_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 21), self.Parent_Material_Accumulation_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 22), self.Parent_Material_Lithology_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 23), self.Parent_Material_Lithology_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 24), self.Drainage_Class_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 25), self.Permeability_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 26), self.Runoff_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 27), self.Erosion_Type_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 28), self.Erosion_Degree_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 29), self.Effective_Soil_Depth_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 30), self.Land_Cover_Main_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 31), self.Land_Cover_Climate_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 32), self.Land_Cover_Vegetation_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 33), self.Land_Cover_Performance_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 34), self.Land_Cover_Main_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 35), self.Land_Cover_Climate_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 36), self.Land_Cover_Vegetation_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 37), self.Land_Cover_Performance_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 38), self.Soil_Moisture_Regime_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 39), self.Soil_Temperature_Regime_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 40), self.Epipedon_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 41), self.Upper_Limit_1_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 42), self.Lower_Limit_1_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 43), self.Sub_Surface_Horizon_1_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 44), self.Upper_Limit_2_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 45), self.Lower_Limit_2_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 46), self.Sub_Surface_Horizon_2_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 47), self.Upper_Limit_3_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 48), self.Lower_Limit_3_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 49), self.Sub_Surface_Horizon_3_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 50), self.Upper_Limit_4_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 51), self.Lower_Limit_4_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 52), self.Sub_Surface_Horizon_4_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 53), self.Upper_Limit_5_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 54), self.Lower_Limit_5_LE.text())
        self.siteModel.setData(self.siteModel.index(row, 55), self.Taxonomy_Year_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 56), self.Great_Group_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 57), self.Sub_Group_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 58), self.PSC_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 59), self.Mineral_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 60), self.Reaction_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 61), self.Temperature_CB.currentText())
        self.siteModel.setData(self.siteModel.index(row, 63), QDateTime.currentDateTime().toString("hh:mm:ss dd/MM/yyyy"))

        self.siteModel.submitAll()
        # set the insert buttton to disabled after updating the data
        self.Insert_PB.setEnabled(False)
        # set the update buttton to enabled after updating the data
        self.Update_PB.setEnabled(True)

        # update the siteMapper index so that after click the update button the fields is not empty
        self.siteMapper.setCurrentIndex(row)
        # Message box of updating data
        updateMessageBox = QMessageBox()
        updateMessageBox.setText("Site Data updated ")
        updateMessageBox.setWindowTitle("Update Site Data")
        updateMessageBox.setStandardButtons(QMessageBox.Ok)
        updateMessageBox.exec()

    def insertData(self):
        """ Inserting the data entered to model and database """
        logging.info("Insert Database button clicked")

        # get the row index from the number form combo box
        row = self.Number_Form_CB.findText(self.Number_Form_CB.currentText())

        ### set the value from each fields to diferent column in site database table
        self.siteRecord.setValue("NoForm", self.Number_Form_CB.currentText())
        self.siteRecord.setValue("SPT", self.SPT_CB.currentText())
        self.siteRecord.setValue("Provinsi", self.Provinsi_CB.currentText())
        self.siteRecord.setValue("Kabupaten", self.Kabupaten_CB.currentText())
        self.siteRecord.setValue("Kecamatan", self.Kecamatan_CB.currentText())
        self.siteRecord.setValue("Desa", self.Desa_CB.currentText())
        self.siteRecord.setValue("Dates", str(self.Date_CB.date().toString("dd/MM/yyyy")))
        self.siteRecord.setValue("InitialName", self.Initial_Name_LE.text())
        self.siteRecord.setValue("ObservationNumber", self.Observation_Number_LE.text())
        self.siteRecord.setValue("KindObservation", self.Kind_Observation_CB.currentText())
        self.siteRecord.setValue("UTMZone1", self.UTM_Zone_1_LE.text())
        self.siteRecord.setValue("UTMZone2", self.UTM_Zone_2_LE.text())
        self.siteRecord.setValue("XEast", self.X_East_LE.text())
        self.siteRecord.setValue("YNorth", self.Y_North_LE.text())
        self.siteRecord.setValue("Elevation", self.Elevation_LE.text())
        self.siteRecord.setValue("Landform", self.Landform_CB.currentText())
        self.siteRecord.setValue("Relief", self.Relief_CB.currentText())
        self.siteRecord.setValue("SlopePercentage", self.Slope_Percentage_LE.text() )
        self.siteRecord.setValue("SlopeShape", self.Slope_Shape_CB.currentText())
        self.siteRecord.setValue("SlopePosition", self.Slope_Position_CB.currentText())
        self.siteRecord.setValue("ModeOfAccumulation", self.Parent_Material_Accumulation_CB.currentText())
        self.siteRecord.setValue("Lithology1", self.Parent_Material_Lithology_1_CB.currentText())
        self.siteRecord.setValue("Lithology2", self.Parent_Material_Lithology_2_CB.currentText())
        self.siteRecord.setValue("DrainageClass", self.Drainage_Class_CB.currentText())
        self.siteRecord.setValue("Permeability", self.Permeability_CB.currentText())
        self.siteRecord.setValue("Runoff", self.Runoff_CB.currentText())
        self.siteRecord.setValue("ErosionType", self.Erosion_Type_CB.currentText())
        self.siteRecord.setValue("ErosionDegree", self.Erosion_Degree_CB.currentText())
        self.siteRecord.setValue("EffectiveSoilDepth", self.Effective_Soil_Depth_LE.text())
        self.siteRecord.setValue("LandCoverMain1", self.Land_Cover_Main_1_CB.currentText())
        self.siteRecord.setValue("LandCoverClimate1", self.Land_Cover_Climate_1_CB.currentText())
        self.siteRecord.setValue("LandCoverVegetation1", self.Land_Cover_Vegetation_1_CB.currentText())
        self.siteRecord.setValue("LandCoverPerformance1", self.Land_Cover_Performance_1_CB.currentText())
        self.siteRecord.setValue("LandCoverMain2", self.Land_Cover_Main_2_CB.currentText())
        self.siteRecord.setValue("LandCoverClimate2", self.Land_Cover_Climate_2_CB.currentText())
        self.siteRecord.setValue("LandCoverVegetation2", self.Land_Cover_Vegetation_2_CB.currentText())
        self.siteRecord.setValue("LandCoverPerformance2", self.Land_Cover_Performance_2_CB.currentText())
        self.siteRecord.setValue("SoilMoistureRegime", self.Soil_Moisture_Regime_CB.currentText())
        self.siteRecord.setValue("SoilTemperatureRegime", self.Soil_Temperature_Regime_CB.currentText())
        self.siteRecord.setValue("Epipedon", self.Epipedon_CB.currentText())
        self.siteRecord.setValue("UpperLimit1", self.Upper_Limit_1_LE.text())
        self.siteRecord.setValue("LowerLimit1", self.Lower_Limit_1_LE.text())
        self.siteRecord.setValue("SubSurfaceHorizon1", self.Sub_Surface_Horizon_1_CB.currentText())
        self.siteRecord.setValue("UpperLimit2", self.Upper_Limit_2_LE.text())
        self.siteRecord.setValue("LowerLimit2", self.Lower_Limit_2_LE.text())
        self.siteRecord.setValue("SubSurfaceHorizon2", self.Sub_Surface_Horizon_2_CB.currentText())
        self.siteRecord.setValue("UpperLimit3", self.Upper_Limit_3_LE.text())
        self.siteRecord.setValue("LowerLimit3", self.Lower_Limit_3_LE.text())
        self.siteRecord.setValue("SubSurfaceHorizon3", self.Sub_Surface_Horizon_3_CB.currentText())
        self.siteRecord.setValue("UpperLimit4", self.Upper_Limit_4_LE.text())
        self.siteRecord.setValue("LowerLimit4", self.Lower_Limit_4_LE.text())
        self.siteRecord.setValue("SubSurfaceHorizon4", self.Sub_Surface_Horizon_4_CB.currentText())
        self.siteRecord.setValue("UpperLimit5", self.Upper_Limit_5_LE.text())
        self.siteRecord.setValue("LowerLimit5", self.Lower_Limit_5_LE.text())
        self.siteRecord.setValue("TaxonomyYear", self.Taxonomy_Year_CB.currentText())
        self.siteRecord.setValue("GreatGroup", self.Great_Group_CB.currentText())
        self.siteRecord.setValue("SubGroup", self.Sub_Group_CB.currentText())
        self.siteRecord.setValue("PSC", self.PSC_CB.currentText())
        self.siteRecord.setValue("Mineral", self.Mineral_CB.currentText())
        self.siteRecord.setValue("Reaction", self.Reaction_CB.currentText())
        self.siteRecord.setValue("Temperature", self.Temperature_CB.currentText())
        self.siteRecord.setValue("Entry_Time", str(QDateTime.currentDateTime().toString("hh:mm:ss dd/MM/yyyy")))

        # if the current index != -1 or not empty submit the data to database and model
        if self.siteModel.insertRecord(-1, self.siteRecord):
            self.siteModel.submitAll()
            
        # disabled the insert button
        self.Insert_PB.setEnabled(False)
        # enabled the update button
        self.Update_PB.setEnabled(True)

        # update the siteMapper index so that after click the update button the fields is not empty
        
        self.siteMapper.toLast()

        # Message box for inserting data
        insertMessageBox = QMessageBox()
        insertMessageBox.setText("Data inserted")
        insertMessageBox.setWindowTitle("Insert Data")
        insertMessageBox.setStandardButtons(QMessageBox.Ok)
        insertMessageBox.exec()

        # set focus to number form
        self.Number_Form_CB.lineEdit().setFocus()
        self.Number_Form_CB.lineEdit().selectAll()

    def numberFormTextChanged(self):
        """ Event that occurs when the number form value/text is changed """

        print("Clik Number Form")
        logging.info("Number Form Clicked/Focused")

        # get the current number/value/text from number form comboBox
        numberFormText = self.Number_Form_CB.currentText()
        # find the index of the current value  
        checkNumberFormText = self.Number_Form_CB.findText(numberFormText)

        # check if number form value is not empty
        if (checkNumberFormText != -1):
            
            ### get the data from the site model to entered in each coresponding fields
            self.SPT_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,2))))
            self.Provinsi_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,3))))
            self.Kabupaten_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,4))))
            self.Kecamatan_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,5))))
            self.Desa_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,6))))
            self.Date_CB.setDate(QDate.fromString(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,7))), "dd/MM/yyyy"))
            self.Initial_Name_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,8))))
            self.Observation_Number_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,9))))
            self.Kind_Observation_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,10))))
            self.UTM_Zone_1_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,11))))
            self.UTM_Zone_2_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,12))))
            self.X_East_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,13))))
            self.Y_North_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,14))))
            self.Elevation_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,15))))
            self.Landform_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,16))))
            self.Relief_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,17))))
            self.Slope_Percentage_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,18))))
            self.Slope_Shape_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,19))))
            self.Slope_Position_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,20))))
            self.Parent_Material_Accumulation_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,21))))
            self.Parent_Material_Lithology_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,22))))
            self.Parent_Material_Lithology_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,23))))
            self.Drainage_Class_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,24))))
            self.Permeability_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,25))))
            self.Runoff_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,26))))
            self.Erosion_Type_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,27))))
            self.Erosion_Degree_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,28))))
            self.Effective_Soil_Depth_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,29))))
            self.Land_Cover_Main_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,30))))
            self.Land_Cover_Climate_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,31))))
            self.Land_Cover_Vegetation_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,32))))
            self.Land_Cover_Performance_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,33))))
            self.Land_Cover_Main_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,34))))
            self.Land_Cover_Climate_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,35))))
            self.Land_Cover_Vegetation_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,36))))
            self.Land_Cover_Performance_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,37))))
            self.Soil_Moisture_Regime_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,38))))
            self.Soil_Temperature_Regime_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,39))))
            self.Epipedon_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,40))))
            self.Upper_Limit_1_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,41))))
            self.Lower_Limit_1_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,42))))
            self.Sub_Surface_Horizon_1_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,43))))
            self.Upper_Limit_2_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,44))))
            self.Lower_Limit_2_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,45))))
            self.Sub_Surface_Horizon_2_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,46))))
            self.Upper_Limit_3_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,47))))
            self.Lower_Limit_3_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,48))))
            self.Sub_Surface_Horizon_3_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,49))))
            self.Upper_Limit_4_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,50))))
            self.Lower_Limit_4_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,51))))
            self.Sub_Surface_Horizon_4_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,52))))
            self.Upper_Limit_5_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,53))))
            self.Lower_Limit_5_LE.setText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,54))))
            self.Taxonomy_Year_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,55))))
            self.Great_Group_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,56))))
            self.Sub_Group_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,57))))
            self.PSC_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,58))))
            self.Mineral_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,59))))
            self.Reaction_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,60))))
            self.Temperature_CB.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNumberFormText,61))))

            # set the current index of mapper to selected number form in the comboBox
            self.siteMapper.setCurrentIndex(checkNumberFormText)

            ### disabled and enabled the buton
            self.Insert_PB.setEnabled(False)
            self.Update_PB.setEnabled(True)
            self.Delete_PB.setEnabled(True)
            self.Copy_PB.setEnabled(False)
            self.Copy_CB.setEnabled(False)
            ## TODO: fixed this add family from english
            self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
            self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
            self.pscEnglish = proxyModelRowIndex(self.pscModel, self.PSC_CB.currentText())
            self.mineralEnglish = proxyModelRowIndex(self.mineralModel, self.Mineral_CB.currentText())
            self.temperatureEnglish = proxyModelRowIndex(self.temperatureModel, self.Temperature_CB.currentText())
            self.Pedon_Label_L.setText(self.subGroupEnglish + " "+ self.greatGroupEnglish + ", " + self.pscEnglish + ", " + self.mineralEnglish +", "+ self.temperatureEnglish)

        else:
            # enabled the copy button and comboBox if the number form is not on the database
            self.Copy_PB.setEnabled(True)
            self.Copy_CB.setEnabled(True)
            # self.Copy_CB.setCurrentText(str(int(self.Number_Form_CB.currentText())-1))
            # set the value of copy comboBox to the last number form index in the siteModel
            self.Copy_CB.setCurrentIndex(self.proxySiteModel.rowCount()-1)

            ### set all the fields in comboBox and lineEdit to empty
            self.SPT_CB.setCurrentIndex(-1)
            self.Provinsi_CB.setCurrentIndex(-1)
            self.Kabupaten_CB.setCurrentIndex(-1)
            self.Kecamatan_CB.setCurrentIndex(-1)
            self.Desa_CB.setCurrentIndex(-1)
            # TODO: set the Date_CB maybe to the 1 january 2019/2020
            # self.Date_CB.setCurrentIndex(-1)
            self.Initial_Name_LE.setText("")
            self.Observation_Number_LE.setText("")
            self.Kind_Observation_CB.setCurrentIndex(-1)
            self.UTM_Zone_1_LE.setText("")
            self.UTM_Zone_2_LE.setText("")
            self.X_East_LE.setText("")
            self.Y_North_LE.setText("")
            self.Elevation_LE.setText("")
            self.Landform_CB.setCurrentIndex(-1)
            self.Relief_CB.setCurrentIndex(-1)
            self.Slope_Percentage_LE.setText("")
            self.Slope_Shape_CB.setCurrentIndex(-1)
            self.Slope_Position_CB.setCurrentIndex(-1)
            self.Parent_Material_Accumulation_CB.setCurrentIndex(-1)
            self.Parent_Material_Lithology_1_CB.setCurrentIndex(-1)
            self.Parent_Material_Lithology_2_CB.setCurrentIndex(-1)
            self.Drainage_Class_CB.setCurrentIndex(-1)
            self.Permeability_CB.setCurrentIndex(-1)
            self.Runoff_CB.setCurrentIndex(-1)
            self.Erosion_Type_CB.setCurrentIndex(-1)
            self.Erosion_Degree_CB.setCurrentIndex(-1)
            self.Effective_Soil_Depth_LE.setText("")
            self.Land_Cover_Main_1_CB.setCurrentIndex(-1)
            self.Land_Cover_Climate_1_CB.setCurrentIndex(-1)
            self.Land_Cover_Vegetation_1_CB.setCurrentIndex(-1)
            self.Land_Cover_Performance_1_CB.setCurrentIndex(-1)
            self.Land_Cover_Main_2_CB.setCurrentIndex(-1)
            self.Land_Cover_Climate_2_CB.setCurrentIndex(-1)
            self.Land_Cover_Vegetation_2_CB.setCurrentIndex(-1)
            self.Land_Cover_Performance_2_CB.setCurrentIndex(-1)
            self.Soil_Moisture_Regime_CB.setCurrentIndex(-1)
            self.Soil_Temperature_Regime_CB.setCurrentIndex(-1)
            self.Epipedon_CB.setCurrentIndex(-1)
            self.Upper_Limit_1_LE.setText("")
            self.Lower_Limit_1_LE.setText("")
            self.Sub_Surface_Horizon_1_CB.setCurrentIndex(-1)
            self.Upper_Limit_2_LE.setText("")
            self.Lower_Limit_2_LE.setText("")
            self.Sub_Surface_Horizon_2_CB.setCurrentIndex(-1)
            self.Upper_Limit_3_LE.setText("")
            self.Lower_Limit_3_LE.setText("")
            self.Sub_Surface_Horizon_3_CB.setCurrentIndex(-1)
            self.Upper_Limit_4_LE.setText("")
            self.Lower_Limit_4_LE.setText("")
            self.Sub_Surface_Horizon_4_CB.setCurrentIndex(-1)
            self.Upper_Limit_5_LE.setText("")
            self.Lower_Limit_5_LE.setText("")
            self.Taxonomy_Year_CB.setCurrentIndex(-1)
            self.Great_Group_CB.setCurrentIndex(-1)
            self.Sub_Group_CB.setCurrentIndex(-1)
            self.PSC_CB.setCurrentIndex(-1)
            self.Mineral_CB.setCurrentIndex(-1)
            self.Reaction_CB.setCurrentIndex(-1)
            self.Temperature_CB.setCurrentIndex(-1)

            # enabled the insert button to true if new number form is entered
            self.Insert_PB.setEnabled(True)
            # disable the update button
            self.Update_PB.setEnabled(False)
            self.Delete_PB.setEnabled(False)
            self.Pedon_Label_L.setText("")

            ### Setting the default Value of the parameter
            # Set UTM Zone
            self.UTM_Zone_1_LE.setText("49")
            self.UTM_Zone_2_LE.setText("S")
            # set the default value to udic "ud"
            self.Soil_Moisture_Regime_CB.lineEdit().setText("ud")
            # set the default value to isohyperthermic "ih"
            self.Soil_Temperature_Regime_CB.lineEdit().setText("ih")
            # set the Taxonomy year
            self.Taxonomy_Year_CB.lineEdit().setText("2014")

    def greatGroupFocusLost(self):
        try:
            self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
            print("Great Group Focus Out")
            self.Pedon_Label_L.setText(self.greatGroupEnglish)
            print("greatGroupEnglish: ", self.greatGroupEnglish)
        except TypeError as te:
            print(te)

    def subGroupFocusLost(self):
        try:
            print("Sub Group Focus Out")
            self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
            self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
            self.Pedon_Label_L.setText(self.subGroupEnglish + " "+self.greatGroupEnglish+", ")
        except TypeError as te:
            print(te)

    def pscFocusLost(self):
        print("PSC Focus Out")
        try:
            self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
            self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
            self.pscEnglish = proxyModelRowIndex(self.pscModel, self.PSC_CB.currentText())
            self.Pedon_Label_L.setText(self.subGroupEnglish+ " "+self.greatGroupEnglish+", "+ self.pscEnglish +", ")
        except TypeError as te:
            print(te)

    def mineralFocusLost(self):
        try:
            print("Mineral Focus Out")
            self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
            self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
            self.pscEnglish = proxyModelRowIndex(self.pscModel, self.PSC_CB.currentText())
            self.mineralEnglish = proxyModelRowIndex(self.mineralModel, self.Mineral_CB.currentText())
            self.Pedon_Label_L.setText(self.subGroupEnglish+ " "+self.greatGroupEnglish+", "+ self.pscEnglish +", "+ self.mineralEnglish +", ")
        except TypeError as te:
            print(te)

    def temperatureFocusLost(self):
        print("Temperature Focus Out")
        self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
        self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
        self.pscEnglish = proxyModelRowIndex(self.pscModel, self.PSC_CB.currentText())
        self.mineralEnglish = proxyModelRowIndex(self.mineralModel, self.Mineral_CB.currentText())
        self.temperatureEnglish = proxyModelRowIndex(self.temperatureModel, self.Temperature_CB.currentText())
        self.Pedon_Label_L.setText(self.subGroupEnglish + " "+ self.greatGroupEnglish + ", " + self.pscEnglish + ", " + self.mineralEnglish +", "+ self.temperatureEnglish)

    def setComboBoxStyle(self):
        """ set the comboBox style to expand the width of the drop-down table. """
        logging.info("Site Windows - Set Combo Box Style")

        ### set all the style using setStyleSheet model
        self.SPT_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Provinsi_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Kabupaten_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Kecamatan_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Desa_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Kind_Observation_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Landform_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:250px}")
        self.Relief_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:180px}")
        self.Slope_Shape_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:150px}")
        self.Slope_Position_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Parent_Material_Accumulation_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:150px}")
        self.Parent_Material_Lithology_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:230px}")
        self.Parent_Material_Lithology_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:230px}")
        self.Drainage_Class_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Permeability_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Runoff_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Erosion_Type_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Erosion_Degree_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Soil_Moisture_Regime_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Soil_Temperature_Regime_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:150px}")
        self.Land_Cover_Main_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Land_Cover_Climate_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Land_Cover_Vegetation_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Land_Cover_Performance_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:180px}")
        self.Land_Cover_Main_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Land_Cover_Climate_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Land_Cover_Vegetation_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Land_Cover_Performance_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:180px}")
        self.Epipedon_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Sub_Surface_Horizon_1_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Sub_Surface_Horizon_2_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Sub_Surface_Horizon_3_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Sub_Surface_Horizon_4_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        # self.Taxonomy_Year_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Great_Group_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:170px}")
        self.Sub_Group_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.PSC_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:200px}")
        self.Mineral_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:170px}")
        self.Reaction_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:180px}")
        self.Temperature_CB.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:150px}")

    def populateData(self):
        """ Populate the data required for the fields """
        logging.info("Site Windows - Populate Data")

        self.Number_Form_CB.setInsertPolicy(self.Number_Form_CB.NoInsert)

        # Number Form
        populate(self.siteModel, "Site", self.Number_Form_CB, self.noFormCompleter)
        self.Number_Form_CB.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))

        # Copy Combobox
        self.proxySiteModel = QSortFilterProxyModel()
        self.proxySiteModel.setSourceModel(self.siteModel)
        # populate(self.proxySiteModel, "Site", self.Copy_CB, self.noFormCompleter)
        self.Copy_CB.setModel(self.proxySiteModel)
        self.Copy_CB.setModelColumn(1)
        self.Copy_CB.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))
        self.Copy_PB.setEnabled(False)
        self.Copy_CB.setEnabled(False)

        # Provinsi
        self.provinsiModel = QSqlQueryModel()
        self.provinsiTableView = QTableView()
        self.provinsiModel.setQuery("select distinct provinsi from location")
        self.Provinsi_CB.setModel(self.provinsiModel)
        self.Provinsi_CB.setCurrentIndex(-1)
        self.Provinsi_CB.setCompleter(self.provinsiCompleter)
        self.provinsiCompleter.setModel(self.provinsiModel)
        # set the provinsi completer to caseinsensitif
        self.provinsiCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        pop = self.provinsiCompleter.popup()
        padding = pop.width() - pop.viewport().width()
        self.provinsiCompleter.popup().setMinimumWidth(200)
        # # set the provinsi to "Kalimantan Tengah"
        # self.Provinsi_CB.lineEdit().setText("Kalimantan Tengah")
       
        # Kabupaten, kecamatan, desa/kompartemen populate using combobox activated method

        # Kind of observation
        self.kindObsModel = QSqlTableModel()
        self.kindObsTableView = QTableView()
        populatePlusTable(self.kindObsModel,"KindObs", self.Kind_Observation_CB, self.kindObservationCompleter, self.kindObsTableView, 80)
        # Set Kind of Observation Completer to Case Insensitive 
        self.kindObservationCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        # kindObsModel.match()

        # Landform
        self.landFormModel = QSqlTableModel()
        self.landFormTableView = QTableView()
        populatePlusTable(self.landFormModel, "Landform", self.Landform_CB, self.landFormCompleter, self.landFormTableView, 200)
        self.Landform_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set Landform completer to Case Insensitive
        self.landFormCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Relief
        self.reliefModel = QSqlTableModel()
        self.reliefTableView = QTableView()
        populatePlusTable(self.reliefModel, "Relief", self.Relief_CB, self.reliefCompleter, self.reliefTableView, 130)
        self.Relief_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Relief Completer to Case Insensitive 
        self.reliefCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Slope Shape
        self.slopeShapeModel = QSqlTableModel()
        self.slopeShapeTableView = QTableView()
        populatePlusTable(self.slopeShapeModel, "SlopeShape", self.Slope_Shape_CB, self.slopeShapeCompleter, self.slopeShapeTableView, 100)
        self.Slope_Shape_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Slope Shape Completer to Case Insensitive 
        self.slopeShapeCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Slope Position
        self.slopePositionModel = QSqlTableModel()
        self.slopePositionTableView = QTableView()
        populatePlusTable(self.slopePositionModel, "SlopePosition", self.Slope_Position_CB, self.slopePositionCompleter, self.slopePositionTableView, 80)
        self.Slope_Position_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Slope Position Completer to Case Insensitive 
        self.slopePositionCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Parent Material Mode of Accumulation
        self.parentMaterialAccumulationModel = QSqlTableModel()
        self.parentMaterialAccumulationTableView = QTableView()
        populatePlusTable(self.parentMaterialAccumulationModel, "ParentMaterialAccu", self.Parent_Material_Accumulation_CB, self.parentMaterialAccumulationCompleter, self.parentMaterialAccumulationTableView, 100)
        self.Parent_Material_Accumulation_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Parent Material Completer to Case Insensitive 
        self.parentMaterialAccumulationCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Parent Material Lithologi 1 & 2
        self.parentMaterialLithology1Model = QSqlTableModel()
        self.parentMaterialLithology2Model = QSqlTableModel()
        self.lithologyTableView = QTableView()
        populatePlusTable(self.parentMaterialLithology1Model, "ParentMaterialLitho", self.Parent_Material_Lithology_1_CB, self.parentMaterialLithology1Completer, self.lithologyTableView, 180)
        populatePlusTable(self.parentMaterialLithology2Model, "ParentMaterialLitho", self.Parent_Material_Lithology_2_CB, self.parentMaterialLithology2Completer, self.lithologyTableView, 180)
        self.Parent_Material_Lithology_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Parent_Material_Lithology_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Soil Drainage Class
        self.drainageClassModel = QSqlTableModel()
        self.drainageClassTableView = QTableView()
        populatePlusTable(self.drainageClassModel, "Drainage", self.Drainage_Class_CB, self.drainageClassCompleter, self.drainageClassTableView, 150)
        self.Drainage_Class_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Permeability
        self.permeabilityModel = QSqlTableModel()
        self.permeabilityTableView = QTableView()
        populatePlusTable(self.permeabilityModel, "Permeability", self.Permeability_CB, self.permeabilityCompleter, self.permeabilityTableView, 80)
        self.Permeability_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Runoff
        self.runoffModel = QSqlTableModel()
        self.runoffTableView = QTableView()
        populatePlusTable(self.runoffModel, "Runoff", self.Runoff_CB, self.runoffCompleter, self.runoffTableView, 80)
        self.Runoff_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Erosion Type
        self.erosionTypeModel = QSqlTableModel()
        self.erosionTypeTableView = QTableView()
        populatePlusTable(self.erosionTypeModel, "ErosionType", self.Erosion_Type_CB, self.erosionTypeCompleter, self.erosionTypeTableView, 110)
        self.Erosion_Type_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Erosion Type Completer to Case Insensitive 
        self.erosionTypeCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Erosion Degree
        self.erosionDegreeModel = QSqlTableModel()
        self.erosionDegreeTableView = QTableView()
        populatePlusTable(self.erosionDegreeModel, "ErosionDegree", self.Erosion_Degree_CB, self.erosionDegreeCompleter, self.erosionDegreeTableView, 110)
        self.Erosion_Degree_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Land Cover Main
        self.landCoverMain1Model = QSqlTableModel()
        self.landCoverMain2Model = QSqlTableModel()
        self.landCoverMainTableView = QTableView()
        populatePlusTable(self.landCoverMain1Model, "LandCoverMain", self.Land_Cover_Main_1_CB, self.landCoverMain1Completer, self.landCoverMainTableView, 150)
        populatePlusTable(self.landCoverMain2Model, "LandCoverMain", self.Land_Cover_Main_2_CB, self.landCoverMain2Completer, self.landCoverMainTableView, 150)
        self.Land_Cover_Main_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Land_Cover_Main_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set the land cover main 1 and 2 Completer to Case Insensitive
        self.landCoverMain1Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.landCoverMain2Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Land Cover Climate
        self.landCoverClimate1Model = QSqlTableModel()
        self.landCoverClimate2Model = QSqlTableModel()
        self.landCoverClimateTableView = QTableView()
        populatePlusTable(self.landCoverClimate1Model, "LandCoverClimate", self.Land_Cover_Climate_1_CB, self.landCoverClimate1Completer, self.landCoverClimateTableView, 80)
        populatePlusTable(self.landCoverClimate2Model, "LandCoverClimate", self.Land_Cover_Climate_2_CB, self.landCoverClimate2Completer, self.landCoverClimateTableView, 80)
        self.Land_Cover_Climate_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Land_Cover_Climate_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set the land cover climate 1 and 2 Completer to Case Insensitive
        self.landCoverClimate1Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.landCoverClimate2Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Land Cover Vegetation
        self.landCoverVegetation1Model = QSqlTableModel()
        self.landCoverVegetation2Model = QSqlTableModel()
        self.landCoverVegetationTableView = QTableView()
        populatePlusTable(self.landCoverVegetation1Model, "LandCoverVegetation", self.Land_Cover_Vegetation_1_CB, self.landCoverVegetation1Completer, self.landCoverVegetationTableView, 150)
        populatePlusTable(self.landCoverVegetation2Model, "LandCoverVegetation", self.Land_Cover_Vegetation_2_CB, self.landCoverVegetation2Completer, self.landCoverVegetationTableView, 150)
        self.Land_Cover_Vegetation_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Land_Cover_Vegetation_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set the land cover vegetation 1 and 2 Completer to Case Insensitive
        self.landCoverVegetation1Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.landCoverVegetation2Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Land Cover Performance
        self.landCoverPerformance1Model = QSqlTableModel()
        self.landCoverPerformance2Model = QSqlTableModel()
        self.landCoverPerformanceTableView = QTableView()
        populatePlusTable(self.landCoverPerformance1Model, "LandCoverPerformance", self.Land_Cover_Performance_1_CB, self.landCoverPerformance1Completer, self.landCoverPerformanceTableView, 130)
        populatePlusTable(self.landCoverPerformance2Model, "LandCoverPerformance", self.Land_Cover_Performance_2_CB, self.landCoverPerformance2Completer, self.landCoverPerformanceTableView, 130)
        self.Land_Cover_Performance_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Land_Cover_Performance_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set the land cover performance1 1 and 2 Completer to Case Insensitive
        self.landCoverPerformance1Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.landCoverPerformance2Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Soil Moisture Regime
        self.soilMoistureRegimeModel = QSqlTableModel()
        self.soilMoistureRegimeTableView = QTableView()
        populatePlusTable(self.soilMoistureRegimeModel, "MoistureRegime", self.Soil_Moisture_Regime_CB, self.soilMoistureRegimeCompleter, self.soilMoistureRegimeTableView, 60)
        self.Soil_Moisture_Regime_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set default soil moisture Regime to udik "ud"
        self.Soil_Moisture_Regime_CB.lineEdit().setText("ud")
        # Set Soil Moisture Regime Completer to Case Insensitive    
        self.soilMoistureRegimeCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Soil Tempterature Regime
        self.soilTemperatureRegimeModel = QSqlTableModel()
        self.soilTemperatureRegimeTableVIew = QTableView()
        populatePlusTable(self.soilTemperatureRegimeModel, "TemperatureRegime", self.Soil_Temperature_Regime_CB, self.soilTemperatureRegimeCompleter, self.soilTemperatureRegimeTableVIew, 100)
        self.Soil_Temperature_Regime_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Soil Temperature Regime to isohyperthermik "ih"
        self.Soil_Temperature_Regime_CB.lineEdit().setText("ih")
        # Set Soil Temperature Regime Completer to Case Insensitive
        self.soilTemperatureRegimeCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Epipedon
        self.epipedonModel = QSqlTableModel()
        self.epipedonTableView = QTableView()
        populatePlusTable(self.epipedonModel, "Epipedon", self.Epipedon_CB, self.epipedonCompleter, self.epipedonTableView, 80)
        self.Epipedon_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # Set Soil Epipedon to ochric "oc"
        self.Epipedon_CB.lineEdit().setText("oc")
        # Set Epipedon Completer to Case Insensitive
        self.epipedonCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Sub Surface Horizon
        self.subSurfaceHorizon1Model = QSqlTableModel()
        self.subSurfaceHorizon2Model = QSqlTableModel()
        self.subSurfaceHorizon3Model = QSqlTableModel()
        self.subSurfaceHorizon4Model = QSqlTableModel()
        self.subSurfaceHorizonTableView = QTableView()
        populatePlusTable(self.subSurfaceHorizon1Model, "SubSurfaceHorizon", self.Sub_Surface_Horizon_1_CB, self.subSurfaceHorizon1Completer, self.subSurfaceHorizonTableView, 90)
        populatePlusTable(self.subSurfaceHorizon2Model, "SubSurfaceHorizon", self.Sub_Surface_Horizon_2_CB, self.subSurfaceHorizon2Completer, self.subSurfaceHorizonTableView, 90)
        populatePlusTable(self.subSurfaceHorizon3Model, "SubSurfaceHorizon", self.Sub_Surface_Horizon_3_CB, self.subSurfaceHorizon3Completer, self.subSurfaceHorizonTableView, 90)
        populatePlusTable(self.subSurfaceHorizon4Model, "SubSurfaceHorizon", self.Sub_Surface_Horizon_4_CB, self.subSurfaceHorizon4Completer, self.subSurfaceHorizonTableView, 90)
        self.Sub_Surface_Horizon_1_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Sub_Surface_Horizon_2_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Sub_Surface_Horizon_3_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        self.Sub_Surface_Horizon_4_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Set Subsurface Horizon 1, 2, 3, and 4 Completer to Case Insensitive
        self.subSurfaceHorizon1Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.subSurfaceHorizon2Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.subSurfaceHorizon3Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.subSurfaceHorizon4Completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Taxonomy Year
        # Set Taxonomy Year alignment to right
        self.Taxonomy_Year_CB.lineEdit().setAlignment(Qt.AlignRight)
        # Set default value of Taxonomy Year to 2014
        self.Taxonomy_Year_CB.lineEdit().setText("2014")

        # Great Group
        self.greatGroupModel = QSqlTableModel()
        self.greatGroupTableView = QTableView()
        populatePlusTable(self.greatGroupModel, "GreatGroup", self.Great_Group_CB, self.greatGroupCompleter, self.greatGroupTableView, 120)
        self.Great_Group_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Set Great Group Completer to Case Insensitive
        self.greatGroupCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Sub Group
        self.subGroupModel = QSqlTableModel()
        self.subGroupTableView = QTableView()
        populatePlusTable(self.subGroupModel, "SubGroup", self.Sub_Group_CB, self.subGroupCompleter, self.subGroupTableView, 150)
        self.Sub_Group_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Set Sub Group Completer to Case Insensitive
        self.subGroupCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Particle Size Class (PSC)
        self.pscModel = QSqlTableModel()
        self.pscTableView = QTableView()
        populatePlusTable(self.pscModel, "PSC", self.PSC_CB, self.pscCompleter, self.pscTableView, 150)
        self.PSC_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Mineral
        self.mineralModel = QSqlTableModel()
        self.mineralTableView = QTableView()
        populatePlusTable(self.mineralModel, "Mineral", self.Mineral_CB, self.mineralCompleter, self.mineralTableView, 120)
        self.Mineral_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Reaction
        self.reactionModel = QSqlTableModel()
        self.reactionTableView = QTableView()
        populatePlusTable(self.reactionModel, "Reaction", self.Reaction_CB, self.reactionCompleter, self.reactionTableView, 130)
        self.Reaction_CB.lineEdit().setAlignment(Qt.AlignHCenter)

        # Temperature
        self.temperatureModel = QSqlTableModel()
        self.temperatureTableView = QTableView()
        populatePlusTable(self.temperatureModel, "TemperatureRegime", self.Temperature_CB, self.temperatureCompleter, self.temperatureTableView, 100)
        self.Temperature_CB.lineEdit().setAlignment(Qt.AlignHCenter)
        # set default Temperature Soil to Isohyperthermik "ih"
        self.Temperature_CB.lineEdit().setText("ih") 
        # Set Temperature Completer to Case Insensitive
        self.temperatureCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # set the Insert, Update, Check, Delete Button to Disable to prevent inserting data with zero/blank number form
        self.Insert_PB.setEnabled(False)
        self.Update_PB.setEnabled(False)
        self.Reload_PB.setEnabled(False)
        self.Delete_PB.setEnabled(False)

        ###
        # self.greatGroupEnglish = proxyModelRowIndex(self.greatGroupModel, self.Great_Group_CB.currentText())
        # self.subGroupEnglish = proxyModelRowIndex(self.subGroupModel, self.Sub_Group_CB.currentText())
        # self.pscEnglish = proxyModelRowIndex(self.pscModel, self.PSC_CB.currentText())
        # self.mineralEnglish = proxyModelRowIndex(self.mineralModel, self.Mineral_CB.currentText())
        # self.temperatureEnglish = proxyModelRowIndex(self.temperatureModel, self.Temperature_CB.currentText())

    def initCompleter(self):
        """ initilize all the completer needed for comboBox, not for the lineEdit widgets"""
        logging.info("Site Windows - Initialize Completer")
        self.noFormCompleter = QCompleter()
        self.sptCompleter = QCompleter()
        self.provinsiCompleter = QCompleter()
        self.kabupatenCompleter = QCompleter()
        self.kecamatanCompleter = QCompleter()
        self.desaCompleter = QCompleter()
        self.initialNameCompleter = QCompleter()
        self.observationNumberCompleter = QCompleter()
        self.kindObservationCompleter = QCompleter()
        self.utmZone1Completer = QCompleter()
        self.utmZone2Completer = QCompleter()
        self.xEastCompleter = QCompleter()
        self.yNorthCompleter = QCompleter()
        self.landFormCompleter = QCompleter()
        self.reliefCompleter = QCompleter()
        self.slopeShapeCompleter = QCompleter()
        self.slopePositionCompleter = QCompleter()
        self.parentMaterialAccumulationCompleter = QCompleter()
        self.parentMaterialLithology1Completer = QCompleter()
        self.parentMaterialLithology2Completer = QCompleter()
        self.drainageClassCompleter = QCompleter()
        self.permeabilityCompleter = QCompleter()
        self.runoffCompleter = QCompleter()
        self.erosionTypeCompleter = QCompleter()
        self.erosionDegreeCompleter = QCompleter()
        self.effectiveSoilDepthCompleter = QCompleter()
        self.landCoverMain1Completer = QCompleter()
        self.landCoverClimate1Completer = QCompleter()
        self.landCoverVegetation1Completer = QCompleter()
        self.landCoverPerformance1Completer = QCompleter()
        self.landCoverMain2Completer = QCompleter()
        self.landCoverClimate2Completer = QCompleter()
        self.landCoverVegetation2Completer = QCompleter()
        self.landCoverPerformance2Completer = QCompleter()
        self.soilMoistureRegimeCompleter = QCompleter()
        self.soilTemperatureRegimeCompleter = QCompleter()
        self.epipedonCompleter = QCompleter()
        self.subSurfaceHorizon1Completer = QCompleter()
        self.subSurfaceHorizon2Completer = QCompleter()
        self.subSurfaceHorizon3Completer = QCompleter()
        self.subSurfaceHorizon4Completer = QCompleter()
        self.greatGroupCompleter = QCompleter()
        self.subGroupCompleter = QCompleter()
        self.pscCompleter = QCompleter()
        self.mineralCompleter = QCompleter()
        self.reactionCompleter = QCompleter()
        self.temperatureCompleter = QCompleter()

    def initMapper(self):
        """ initialize the site mapper for easy moving between index"""
        logging.info("Site Windows - Initialize Mapper")
        self.siteMapper = QDataWidgetMapper()
        # set the siteMapper model
        self.siteMapper.setModel(self.siteModel)
        # set submit policy to manual
        self.siteMapper.setSubmitPolicy(QDataWidgetMapper.ManualSubmit)

        ### get the data from model to mapped to comboBox or lineEdit widget
        self.siteMapper.addMapping(self.Number_Form_CB, 1)
        self.siteMapper.addMapping(self.SPT_CB, 2)
        self.siteMapper.addMapping(self.Provinsi_CB, 3)
        self.siteMapper.addMapping(self.Kabupaten_CB, 4)
        self.siteMapper.addMapping(self.Kecamatan_CB, 5)
        self.siteMapper.addMapping(self.Desa_CB, 6)
        self.siteMapper.addMapping(self.Date_CB, 7)
        self.siteMapper.addMapping(self.Initial_Name_LE, 8)
        self.siteMapper.addMapping(self.Observation_Number_LE, 9)
        self.siteMapper.addMapping(self.Kind_Observation_CB, 10)
        self.siteMapper.addMapping(self.UTM_Zone_1_LE, 11)
        self.siteMapper.addMapping(self.UTM_Zone_2_LE, 12)
        self.siteMapper.addMapping(self.X_East_LE, 13)
        self.siteMapper.addMapping(self.Y_North_LE, 14)
        self.siteMapper.addMapping(self.Elevation_LE, 15)
        self.siteMapper.addMapping(self.Landform_CB, 16)
        self.siteMapper.addMapping(self.Relief_CB, 17)
        self.siteMapper.addMapping(self.Slope_Percentage_LE, 18)
        self.siteMapper.addMapping(self.Slope_Shape_CB, 19)
        self.siteMapper.addMapping(self.Slope_Position_CB, 20)
        self.siteMapper.addMapping(self.Parent_Material_Accumulation_CB, 21)
        self.siteMapper.addMapping(self.Parent_Material_Lithology_1_CB, 22)
        self.siteMapper.addMapping(self.Parent_Material_Lithology_2_CB, 23)
        self.siteMapper.addMapping(self.Drainage_Class_CB, 24)
        self.siteMapper.addMapping(self.Permeability_CB, 25)
        self.siteMapper.addMapping(self.Runoff_CB, 26)
        self.siteMapper.addMapping(self.Erosion_Type_CB, 27)
        self.siteMapper.addMapping(self.Erosion_Degree_CB, 28)
        self.siteMapper.addMapping(self.Effective_Soil_Depth_LE, 29)
        self.siteMapper.addMapping(self.Land_Cover_Main_1_CB, 30)
        self.siteMapper.addMapping(self.Land_Cover_Climate_1_CB, 31)
        self.siteMapper.addMapping(self.Land_Cover_Vegetation_1_CB, 32)
        self.siteMapper.addMapping(self.Land_Cover_Performance_1_CB, 33)
        self.siteMapper.addMapping(self.Land_Cover_Main_2_CB, 34)
        self.siteMapper.addMapping(self.Land_Cover_Climate_2_CB, 35)
        self.siteMapper.addMapping(self.Land_Cover_Vegetation_2_CB, 36)
        self.siteMapper.addMapping(self.Land_Cover_Performance_2_CB, 37)
        self.siteMapper.addMapping(self.Soil_Moisture_Regime_CB, 38)
        self.siteMapper.addMapping(self.Soil_Temperature_Regime_CB, 39)
        self.siteMapper.addMapping(self.Epipedon_CB, 40)
        self.siteMapper.addMapping(self.Upper_Limit_1_LE, 41)
        self.siteMapper.addMapping(self.Lower_Limit_1_LE, 42)
        self.siteMapper.addMapping(self.Sub_Surface_Horizon_1_CB, 43)
        self.siteMapper.addMapping(self.Upper_Limit_2_LE, 44)
        self.siteMapper.addMapping(self.Lower_Limit_2_LE, 45)
        self.siteMapper.addMapping(self.Sub_Surface_Horizon_2_CB, 46)
        self.siteMapper.addMapping(self.Upper_Limit_3_LE, 47)
        self.siteMapper.addMapping(self.Lower_Limit_3_LE, 48)
        self.siteMapper.addMapping(self.Sub_Surface_Horizon_3_CB, 49)
        self.siteMapper.addMapping(self.Upper_Limit_4_LE, 50)
        self.siteMapper.addMapping(self.Lower_Limit_4_LE, 51)
        self.siteMapper.addMapping(self.Sub_Surface_Horizon_4_CB, 52)
        self.siteMapper.addMapping(self.Upper_Limit_5_LE, 53)
        self.siteMapper.addMapping(self.Lower_Limit_5_LE, 54)
        self.siteMapper.addMapping(self.Taxonomy_Year_CB, 55)
        self.siteMapper.addMapping(self.Great_Group_CB, 56)
        self.siteMapper.addMapping(self.Sub_Group_CB, 57)
        self.siteMapper.addMapping(self.PSC_CB, 58)
        self.siteMapper.addMapping(self.Mineral_CB, 59)
        self.siteMapper.addMapping(self.Reaction_CB, 60)
        self.siteMapper.addMapping(self.Temperature_CB, 61)

    def kecamatanTextChanged(self):
        """ event that activated Desa combobox - dependant combobox """
        logging.info("Site Windows - Kecamatan ComboBox text Changed")

        # clear the desa model
        self.Desa_CB.model().clear()
        ### init the desa model
        desaModel = QSqlQueryModel()
        query = QSqlQuery()
        query.prepare("Select distinct DESA from Location where KECAMATAN = ?")
        query.addBindValue(self.Kecamatan_CB.currentText())
        query.exec_()
        desaModel.setQuery(query)
        ### set the model, completer for the combobox
        self.Desa_CB.setModel(desaModel)
        self.desaCompleter.setModel(desaModel)
        self.Desa_CB.setCompleter(self.desaCompleter)
        # set the default value of Desa to empty
        self.Desa_CB.setCurrentIndex(-1)

    def kabupatenKotaTextChanged(self):
        """ event that activated Kecamatan combobox - dependant combobox """
        logging.info("Site Windows - Kabupaten ComboBox text Changed")

        # clear kecamatan and desa model
        self.Kecamatan_CB.model().clear()
        self.Desa_CB.model().clear()
        # init the kecamatan model using QSqlQueryModel
        kecamatanModel = QSqlQueryModel()
        query = QSqlQuery()
        query.prepare("Select distinct KECAMATAN from Location where KABKOT = ?")
        query.addBindValue(self.Kabupaten_CB.currentText())
        query.exec_()
        kecamatanModel.setQuery(query)
        # set the model, completer
        self.Kecamatan_CB.setModel(kecamatanModel)
        self.kecamatanCompleter.setModel(kecamatanModel)
        # set kecamatan completer to Case Insensitive
        self.kecamatanCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.Kecamatan_CB.setCompleter(self.kecamatanCompleter)
        ## set the default comboBox value to empty
        self.Kecamatan_CB.setCurrentIndex(-1)
        self.Desa_CB.setCurrentIndex(-1)

    def provinsiTextChanged(self):
        """ event that activated Kabupaten combobox - dependant combobox """
        logging.info("Site Windows - Provinsi ComboBox text Changed")

        ### clear the Kabupaten, Kecamatan, Desa combobox model
        self.Kabupaten_CB.model().clear()
        self.Kecamatan_CB.model().clear()
        self.Desa_CB.model().clear()
        # initialize model for kabupatenKota using QSqlQueryModel
        kabupatenKotaModel = QSqlQueryModel()
        # add QSqlQuery
        query = QSqlQuery()
        query.prepare("Select distinct KABKOT from Location where PROVINSI = ? ")
        query.addBindValue(self.Provinsi_CB.currentText())
        # have to execute the query to works
        query.exec_()
        # set the query for the model
        kabupatenKotaModel.setQuery(query)
        ### set the kabupaten model, completer model and set the completer for Kabupaten ComboBox
        self.Kabupaten_CB.setModel(kabupatenKotaModel)
        self.kabupatenCompleter.setModel(kabupatenKotaModel)
        # Set KabupatenCompleter to CaseInsensitive
        self.kabupatenCompleter.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.Kabupaten_CB.setCompleter(self.kabupatenCompleter)
        ### set the initial value/text (default) to empty
        self.Kabupaten_CB.setCurrentIndex(-1)
        self.Kecamatan_CB.setCurrentIndex(-1)
        self.Desa_CB.setCurrentIndex(-1)

    def lowerLimit1TextChanged(self):
        logging.info("Site Windows - Lower Limit Text Changed")

        # set the sub surface horizon 1 upper limit value the same as the lower of epipedon value
        self.Upper_Limit_2_LE.setText(self.Lower_Limit_1_LE.text())

    def toFirst(self):
        logging.info("Site Windows - Move to the First Index")

        # map to the first index in the number form model in the combobox
        self.siteMapper.toFirst()

    def toPrevious(self):
        logging.info("Site Windows - Move to the Previous Index")

        # map to the previous index in the number form model in the combobox
        self.siteMapper.toPrevious()

    def toNext(self):
        logging.info("Site Windows - Move to the Next Index")

        # map to the next index in the number form model in the combobox
        self.siteMapper.toNext()

    def toLast(self):
        logging.info("Site Windows - Move to the Last Index")

        # map to the last index in the number form model in the combobox
        self.siteMapper.toLast()

    def check(self):
        print("klik")
        # self.Number_Form_CB.model().clear()
       
        # # self.Date_CB.setDat
        # value = "10-06-2018"
        # self.Date_CB.setDate(QDate.fromString(value,"dd-MM-yyyy"))
        # print(str(self.Date_CB.date().toString("dd-MM-yyyy")))

        # # How to get list of Provinsi
        # listProvinsi = []
        # for row in range(provinsiModel.rowCount()):
        #     list = provinsiModel.data(provinsiModel.index(row, 0), Qt.DisplayRole)
        #     listProvinsi.append(list)
        #     print(list)
        # print(listProvinsi)

        # # to get the children/GroupBox of central widget/mainwindow
        # for widgetList in self.centralwidget.children():
        #     # to get QcomboBox from Groupbox
        #     for combo in widgetList.children():
        #         # Check if QComboBox exist
        #         if isinstance(combo, QComboBox):
        #             print(combo.objectName())


            # if isinstance(widgetList, QComboBox):
            #     print("ComboBox: %s " % (widgetList.objectName()))
            #     widgetList.setCurrentText("")
        # print(widgetList)

    def showTable(self):
        """ show the Site Table database in table view """
        logging.info("Site Windows - Show database tables")

        
        # setup the widget ui
        dialog = QDialog()
        vLayout = QVBoxLayout()
        tableView = QTableView()
        tableView.setModel(self.siteModel)
        vLayout.addWidget(tableView)
        dialog.setLayout(vLayout)
        dialog.setWindowTitle("Site Tables")
        dialog.setGeometry(130, 65, 1100, 600)
        dialog.exec()

    def closeEvent(self, event):
        logging.info("Site Entry Closed")

    

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # connect to the database and check if not connected, close the window or maybe throw error
    if not connection.createConnection():
        sys.exit(1)

    main = SiteWindow()
    main.show()

    sys.exit(app.exec_())