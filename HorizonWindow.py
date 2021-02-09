import sys
import logging

from PyQt5.QtWidgets import  (QPushButton, QVBoxLayout, QDialog, QMessageBox, 
                                QTableView, QCompleter, QMainWindow, QLabel, 
                                QApplication, QTextEdit, QDataWidgetMapper)
from PyQt5.QtCore import QDate, QModelIndex, QObject, pyqtSignal, Qt, QEvent, QSortFilterProxyModel, QDateTime
from PyQt5.QtSql import QSqlQuery, QSqlTableModel, QSqlRecord, QSqlQueryModel
from PyQt5 import uic, QtCore

import connection
import sqlite3

import ArrayList as list

uifile = "UI/horizonAll.ui"
Ui_MainWindow, QtBaseClass = uic.loadUiType(uifile)


def munsellChange(hue, valueField, chromaField, valueCompleter, chromaCompleter):
    valueQuery = QSqlQuery()
    valueQuery.prepare('select distinct VALUE from SoilColor where HUE=?')
    valueQuery.addBindValue(hue)
    valueQuery.exec_()
    valueModel = QSqlQueryModel()
    valueModel.setQuery(valueQuery)
    valueField.setModel(valueModel)
    valueField.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))
    valueField.setCurrentIndex(-1)
    valueCompleter.setModel(valueModel)
    valueField.setCompleter(valueCompleter)

    chromaQuery = QSqlQuery()
    chromaQuery.prepare('select distinct CHROMA from SoilColor where HUE=?')
    chromaQuery.addBindValue(hue)
    chromaQuery.exec_()
    chromaModel = QSqlQueryModel()
    chromaModel.setQuery(chromaQuery)
    chromaField.setModel(chromaModel)
    chromaField.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))
    chromaField.setCurrentIndex(-1)
    chromaCompleter.setModel(chromaModel)
    chromaField.setCompleter(chromaCompleter)

def initializeModel(model, table, filter):
    model.setTable(table)

    model.setEditStrategy(QSqlTableModel.OnManualSubmit)
    model.setFilter(filter)
    model.select()

class FocusOutFilter(QObject):
    focusOut = pyqtSignal()

    def eventFilter(self, widget, event):
        if event.type() == QEvent.FocusOut:
            print("--eventFilter() focus_out on: " + widget.objectName())
            self.focusOut.emit()
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

def populate(model, table, combobox):
    # textureClassModel = QSqlTableModel()
    initializeModel(model, table, '')
    combobox.setModel(model)
    combobox.setModelColumn(1)
    combobox.setCurrentIndex(-1)
    # completer.setModel(model)
    # completer.setCompletionColumn(1)
    # combobox.setCompleter(completer)

def populatePlusTable(model, table, combobox, completer, tableView, width2):
    # textureClassModel = QSqlTableModel()
    initializeModel(model, table, '')
    combobox.setModel(model)
    combobox.setModelColumn(0)
    combobox.setCurrentIndex(-1)
    completer.setModel(model)
    # Set Completer Case insensitive
    completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
    # completer.setModelSorting(completer.CaseInsensitivelySortedModel)
    combobox.setCompleter(completer)
    # sorting the model from '0' column
    combobox.model().sort(0, Qt.SortOrder(Qt.AscendingOrder))

    # set the completer table view
    completerTableView = QTableView()
    ## hidden the vertical and horizontal header of the tableview in completer
    completerTableView.verticalHeader().hide()
    completerTableView.horizontalHeader().hide()
    # set the table view with setPopup function
    completer.setPopup(completerTableView)
    # hide TEXT_INDO colum - index = 1
    completerTableView.hideColumn(1)
    # set the 1st column width
    completerTableView.setColumnWidth(0, 40)
    # set the 2nd column width
    completerTableView.setColumnWidth(2, 120)
    # set the width of the completer to beyond the width of the combobox
    completer.popup().setMinimumWidth(165)
    # Hack solution, TODO: must find better solution so the height is dynamically scaled
    completer.popup().setMinimumHeight(50)
    # completer.popup().setMinimumHeight(completerTableView.sizeHintForRow(completer.completionCount()))
    # completer.popup().resizeRowToContents(completer.model().rowCount())
    
    ## initialize TableView for Drop-down combobox
    tableView = QTableView()
    tableView.verticalHeader().hide()
    tableView.horizontalHeader().hide()
    combobox.setView(tableView)
    tableView.hideColumn(1)
    tableView.setColumnWidth(0, 50)
    tableView.setColumnWidth(2, width2)

class HorizonWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        Ui_MainWindow.__init__(self)

        self.horizonModel = QSqlTableModel()
        horizonFilter = ''
        initializeModel(self.horizonModel, 'HorizonAll', horizonFilter)

        self.siteModel = QSqlTableModel()
        initializeModel(self.siteModel, "Site", horizonFilter)

        self.setupUi(self)
        self.initMapper()
        self.initCompleter()

        # Populate Data/Model Horizon
        self.populateDataHor1()
        self.populateDataHor2()
        self.populateDataHor3()
        self.populateDataHor4()
        self.populateDataHor5()
        self.populateDataHor6()

        self.setTableViewCombobox()
        self.setComboboxStyle()
        self.setEvent()

        self.horizonRowCount = self.horizonModel.rowCount

        self.horizonRecord = QSqlRecord()
        self.horizonRecord = self.horizonModel.record()

    def insertData(self):
        print("insert")
        logging.info("Horizon Windows - Insert Button Clicked")

        ### add data using QSqlRecord
        ### self.horizonRecord.setValue("id", "4")

        self.horizonRecord.setValue("NoForm", self.Number_Form_CB_Hor.currentText())

        ### Horizon 1
        self.horizonRecord.setValue("HorDesignDiscon_Hor1", self.Hor_Design_Discon_LE_Hor_1.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor1", self.Hor_Design_Master_LE_Hor_1.text())
        self.horizonRecord.setValue("HorDesignSub_Hor1", self.Hor_Design_Sub_LE_Hor_1.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor1", self.Hor_Design_Number_LE_Hor_1.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor1", self.Hor_Upper_From_LE_Hor_1.text())
        self.horizonRecord.setValue("HorUpperTo_Hor1", self.Hor_Upper_To_LE_Hor_1.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor1", self.Hor_Lower_From_LE_Hor_1.text())
        self.horizonRecord.setValue("HorLowerTo_Hor1", self.Hor_Lower_To_LE_Hor_1.text())
        self.horizonRecord.setValue("BoundaryDist_Hor1", self.Boundary_Dist_CB_Hor_1.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor1", self.Boundary_Topo_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor1", self.Soil_Color_Hue_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor1", self.Soil_Color_Value_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor1", self.Soil_Color_Chroma_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor1", self.Soil_Color_Hue_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor1", self.Soil_Color_Value_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor1", self.Soil_Color_Chroma_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor1", self.Soil_Color_Hue_3_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor1", self.Soil_Color_Value_3_CB_Hor_1.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor1", self.Soil_Color_Chroma_3_CB_Hor_1.currentText())
        self.horizonRecord.setValue("TextureClass_Hor1", self.Texture_Class_CB_Hor_1.currentText())
        self.horizonRecord.setValue("TextureSand_Hor1", self.Texture_Sand_CB_Hor_1.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor1", self.Texture_Mod_Size_CB_Hor_1.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor1", self.Texture_Mod_Abundance_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor1", self.Structure_Shape_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor1", self.Structure_Size_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor1", self.Structure_Grade_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor1", self.Structure_Relation_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor1", self.Structure_Shape_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor1", self.Structure_Size_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor1", self.Structure_Grade_2_CB_Hor_1.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor1", self.Consistence_Moist_CB_Hor_1.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor1", self.Consistence_Stickiness_CB_Hor_1.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor1", self.Consistence_Plasticity_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor1", self.Mottle_Abundance_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor1", self.Mottle_Size_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor1", self.Mottle_Contrast_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor1", self.Mottle_Shape_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor1", self.Mottle_Hue_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor1", self.Mottle_Value_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor1", self.Mottle_Chroma_1_CB_Hor_1.currentText())
        self.horizonRecord.setValue("RootFine_Hor1", self.Root_Fine_CB_Hor_1.currentText())
        self.horizonRecord.setValue("RootMedium_Hor1", self.Root_Medium_CB_Hor_1.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor1", self.Root_Coarse_CB_Hor_1.currentText())

        ### Horizon 2
        self.horizonRecord.setValue("HorDesignDiscon_Hor2", self.Hor_Design_Discon_LE_Hor_2.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor2", self.Hor_Design_Master_LE_Hor_2.text())
        self.horizonRecord.setValue("HorDesignSub_Hor2", self.Hor_Design_Sub_LE_Hor_2.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor2", self.Hor_Design_Number_LE_Hor_2.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor2", self.Hor_Upper_From_LE_Hor_2.text())
        self.horizonRecord.setValue("HorUpperTo_Hor2", self.Hor_Upper_To_LE_Hor_2.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor2", self.Hor_Lower_From_LE_Hor_2.text())
        self.horizonRecord.setValue("HorLowerTo_Hor2", self.Hor_Lower_To_LE_Hor_2.text())
        self.horizonRecord.setValue("BoundaryDist_Hor2", self.Boundary_Dist_CB_Hor_2.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor2", self.Boundary_Topo_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor2", self.Soil_Color_Hue_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor2", self.Soil_Color_Value_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor2", self.Soil_Color_Chroma_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor2", self.Soil_Color_Hue_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor2", self.Soil_Color_Value_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor2", self.Soil_Color_Chroma_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor2", self.Soil_Color_Hue_3_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor2", self.Soil_Color_Value_3_CB_Hor_2.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor2", self.Soil_Color_Chroma_3_CB_Hor_2.currentText())
        self.horizonRecord.setValue("TextureClass_Hor2", self.Texture_Class_CB_Hor_2.currentText())
        self.horizonRecord.setValue("TextureSand_Hor2", self.Texture_Sand_CB_Hor_2.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor2", self.Texture_Mod_Size_CB_Hor_2.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor2", self.Texture_Mod_Abundance_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor2", self.Structure_Shape_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor2", self.Structure_Size_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor2", self.Structure_Grade_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor2", self.Structure_Relation_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor2", self.Structure_Shape_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor2", self.Structure_Size_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor2", self.Structure_Grade_2_CB_Hor_2.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor2", self.Consistence_Moist_CB_Hor_2.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor2", self.Consistence_Stickiness_CB_Hor_2.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor2", self.Consistence_Plasticity_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor2", self.Mottle_Abundance_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor2", self.Mottle_Size_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor2", self.Mottle_Contrast_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor2", self.Mottle_Shape_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor2", self.Mottle_Hue_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor2", self.Mottle_Value_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor2", self.Mottle_Chroma_1_CB_Hor_2.currentText())
        self.horizonRecord.setValue("RootFine_Hor2", self.Root_Fine_CB_Hor_2.currentText())
        self.horizonRecord.setValue("RootMedium_Hor2", self.Root_Medium_CB_Hor_2.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor2", self.Root_Coarse_CB_Hor_2.currentText())

        ### Horizon 3
        self.horizonRecord.setValue("HorDesignDiscon_Hor3", self.Hor_Design_Discon_LE_Hor_3.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor3", self.Hor_Design_Master_LE_Hor_3.text())
        self.horizonRecord.setValue("HorDesignSub_Hor3", self.Hor_Design_Sub_LE_Hor_3.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor3", self.Hor_Design_Number_LE_Hor_3.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor3", self.Hor_Upper_From_LE_Hor_3.text())
        self.horizonRecord.setValue("HorUpperTo_Hor3", self.Hor_Upper_To_LE_Hor_3.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor3", self.Hor_Lower_From_LE_Hor_3.text())
        self.horizonRecord.setValue("HorLowerTo_Hor3", self.Hor_Lower_To_LE_Hor_3.text())
        self.horizonRecord.setValue("BoundaryDist_Hor3", self.Boundary_Dist_CB_Hor_3.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor3", self.Boundary_Topo_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor3", self.Soil_Color_Hue_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor3", self.Soil_Color_Value_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor3", self.Soil_Color_Chroma_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor3", self.Soil_Color_Hue_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor3", self.Soil_Color_Value_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor3", self.Soil_Color_Chroma_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor3", self.Soil_Color_Hue_3_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor3", self.Soil_Color_Value_3_CB_Hor_3.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor3", self.Soil_Color_Chroma_3_CB_Hor_3.currentText())
        self.horizonRecord.setValue("TextureClass_Hor3", self.Texture_Class_CB_Hor_3.currentText())
        self.horizonRecord.setValue("TextureSand_Hor3", self.Texture_Sand_CB_Hor_3.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor3", self.Texture_Mod_Size_CB_Hor_3.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor3", self.Texture_Mod_Abundance_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor3", self.Structure_Shape_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor3", self.Structure_Size_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor3", self.Structure_Grade_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor3", self.Structure_Relation_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor3", self.Structure_Shape_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor3", self.Structure_Size_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor3", self.Structure_Grade_2_CB_Hor_3.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor3", self.Consistence_Moist_CB_Hor_3.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor3", self.Consistence_Stickiness_CB_Hor_3.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor3", self.Consistence_Plasticity_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor3", self.Mottle_Abundance_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor3", self.Mottle_Size_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor3", self.Mottle_Contrast_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor3", self.Mottle_Shape_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor3", self.Mottle_Hue_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor3", self.Mottle_Value_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor3", self.Mottle_Chroma_1_CB_Hor_3.currentText())
        self.horizonRecord.setValue("RootFine_Hor3", self.Root_Fine_CB_Hor_3.currentText())
        self.horizonRecord.setValue("RootMedium_Hor3", self.Root_Medium_CB_Hor_3.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor3", self.Root_Coarse_CB_Hor_3.currentText())

        ### Horizon 4
        self.horizonRecord.setValue("HorDesignDiscon_Hor4", self.Hor_Design_Discon_LE_Hor_4.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor4", self.Hor_Design_Master_LE_Hor_4.text())
        self.horizonRecord.setValue("HorDesignSub_Hor4", self.Hor_Design_Sub_LE_Hor_4.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor4", self.Hor_Design_Number_LE_Hor_4.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor4", self.Hor_Upper_From_LE_Hor_4.text())
        self.horizonRecord.setValue("HorUpperTo_Hor4", self.Hor_Upper_To_LE_Hor_4.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor4", self.Hor_Lower_From_LE_Hor_4.text())
        self.horizonRecord.setValue("HorLowerTo_Hor4", self.Hor_Lower_To_LE_Hor_4.text())
        self.horizonRecord.setValue("BoundaryDist_Hor4", self.Boundary_Dist_CB_Hor_4.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor4", self.Boundary_Topo_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor4", self.Soil_Color_Hue_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor4", self.Soil_Color_Value_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor4", self.Soil_Color_Chroma_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor4", self.Soil_Color_Hue_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor4", self.Soil_Color_Value_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor4", self.Soil_Color_Chroma_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor4", self.Soil_Color_Hue_3_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor4", self.Soil_Color_Value_3_CB_Hor_4.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor4", self.Soil_Color_Chroma_3_CB_Hor_4.currentText())
        self.horizonRecord.setValue("TextureClass_Hor4", self.Texture_Class_CB_Hor_4.currentText())
        self.horizonRecord.setValue("TextureSand_Hor4", self.Texture_Sand_CB_Hor_4.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor4", self.Texture_Mod_Size_CB_Hor_4.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor4", self.Texture_Mod_Abundance_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor4", self.Structure_Shape_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor4", self.Structure_Size_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor4", self.Structure_Grade_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor4", self.Structure_Relation_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor4", self.Structure_Shape_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor4", self.Structure_Size_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor4", self.Structure_Grade_2_CB_Hor_4.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor4", self.Consistence_Moist_CB_Hor_4.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor4", self.Consistence_Stickiness_CB_Hor_4.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor4", self.Consistence_Plasticity_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor4", self.Mottle_Abundance_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor4", self.Mottle_Size_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor4", self.Mottle_Contrast_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor4", self.Mottle_Shape_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor4", self.Mottle_Hue_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor4", self.Mottle_Value_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor4", self.Mottle_Chroma_1_CB_Hor_4.currentText())
        self.horizonRecord.setValue("RootFine_Hor4", self.Root_Fine_CB_Hor_4.currentText())
        self.horizonRecord.setValue("RootMedium_Hor4", self.Root_Medium_CB_Hor_4.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor4", self.Root_Coarse_CB_Hor_4.currentText())

        ### Horizon 5
        self.horizonRecord.setValue("HorDesignDiscon_Hor5", self.Hor_Design_Discon_LE_Hor_5.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor5", self.Hor_Design_Master_LE_Hor_5.text())
        self.horizonRecord.setValue("HorDesignSub_Hor5", self.Hor_Design_Sub_LE_Hor_5.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor5", self.Hor_Design_Number_LE_Hor_5.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor5", self.Hor_Upper_From_LE_Hor_5.text())
        self.horizonRecord.setValue("HorUpperTo_Hor5", self.Hor_Upper_To_LE_Hor_5.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor5", self.Hor_Lower_From_LE_Hor_5.text())
        self.horizonRecord.setValue("HorLowerTo_Hor5", self.Hor_Lower_To_LE_Hor_5.text())
        self.horizonRecord.setValue("BoundaryDist_Hor5", self.Boundary_Dist_CB_Hor_5.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor5", self.Boundary_Topo_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor5", self.Soil_Color_Hue_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor5", self.Soil_Color_Value_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor5", self.Soil_Color_Chroma_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor5", self.Soil_Color_Hue_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor5", self.Soil_Color_Value_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor5", self.Soil_Color_Chroma_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor5", self.Soil_Color_Hue_3_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor5", self.Soil_Color_Value_3_CB_Hor_5.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor5", self.Soil_Color_Chroma_3_CB_Hor_5.currentText())
        self.horizonRecord.setValue("TextureClass_Hor5", self.Texture_Class_CB_Hor_5.currentText())
        self.horizonRecord.setValue("TextureSand_Hor5", self.Texture_Sand_CB_Hor_5.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor5", self.Texture_Mod_Size_CB_Hor_5.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor5", self.Texture_Mod_Abundance_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor5", self.Structure_Shape_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor5", self.Structure_Size_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor5", self.Structure_Grade_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor5", self.Structure_Relation_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor5", self.Structure_Shape_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor5", self.Structure_Size_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor5", self.Structure_Grade_2_CB_Hor_5.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor5", self.Consistence_Moist_CB_Hor_5.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor5", self.Consistence_Stickiness_CB_Hor_5.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor5", self.Consistence_Plasticity_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor5", self.Mottle_Abundance_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor5", self.Mottle_Size_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor5", self.Mottle_Contrast_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor5", self.Mottle_Shape_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor5", self.Mottle_Hue_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor5", self.Mottle_Value_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor5", self.Mottle_Chroma_1_CB_Hor_5.currentText())
        self.horizonRecord.setValue("RootFine_Hor5", self.Root_Fine_CB_Hor_5.currentText())
        self.horizonRecord.setValue("RootMedium_Hor5", self.Root_Medium_CB_Hor_5.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor5", self.Root_Coarse_CB_Hor_5.currentText())

        ### Horizon 6
        self.horizonRecord.setValue("HorDesignDiscon_Hor6", self.Hor_Design_Discon_LE_Hor_6.text())
        self.horizonRecord.setValue("HorDesignMaster_Hor6", self.Hor_Design_Master_LE_Hor_6.text())
        self.horizonRecord.setValue("HorDesignSub_Hor6", self.Hor_Design_Sub_LE_Hor_6.text())
        self.horizonRecord.setValue("HorDesignNumber_Hor6", self.Hor_Design_Number_LE_Hor_6.text())
        self.horizonRecord.setValue("HorUpperFrom_Hor6", self.Hor_Upper_From_LE_Hor_6.text())
        self.horizonRecord.setValue("HorUpperTo_Hor6", self.Hor_Upper_To_LE_Hor_6.text())
        self.horizonRecord.setValue("HorLowerFrom_Hor6", self.Hor_Lower_From_LE_Hor_6.text())
        self.horizonRecord.setValue("HorLowerTo_Hor6", self.Hor_Lower_To_LE_Hor_6.text())
        self.horizonRecord.setValue("BoundaryDist_Hor6", self.Boundary_Dist_CB_Hor_6.currentText())
        self.horizonRecord.setValue("BoundaryTopo_Hor6", self.Boundary_Topo_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorHue1_Hor6", self.Soil_Color_Hue_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorValue1_Hor6", self.Soil_Color_Value_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorChroma1_Hor6", self.Soil_Color_Chroma_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorHue2_Hor6", self.Soil_Color_Hue_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorValue2_Hor6", self.Soil_Color_Value_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorChroma2_Hor6", self.Soil_Color_Chroma_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorHue3_Hor6", self.Soil_Color_Hue_3_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorValue3_Hor6", self.Soil_Color_Value_3_CB_Hor_6.currentText())
        self.horizonRecord.setValue("SoilColorChroma3_Hor6", self.Soil_Color_Chroma_3_CB_Hor_6.currentText())
        self.horizonRecord.setValue("TextureClass_Hor6", self.Texture_Class_CB_Hor_6.currentText())
        self.horizonRecord.setValue("TextureSand_Hor6", self.Texture_Sand_CB_Hor_6.currentText())
        self.horizonRecord.setValue("TextureModSize_Hor6", self.Texture_Mod_Size_CB_Hor_6.currentText())
        self.horizonRecord.setValue("TextureModAbundance_Hor6", self.Texture_Mod_Abundance_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureShape1_Hor6", self.Structure_Shape_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureSize1_Hor6", self.Structure_Size_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureGrade1_Hor6", self.Structure_Grade_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureRelation_Hor6", self.Structure_Relation_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureShape2_Hor6", self.Structure_Shape_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureSize2_Hor6", self.Structure_Size_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("StructureGrade2_Hor6", self.Structure_Grade_2_CB_Hor_6.currentText())
        self.horizonRecord.setValue("ConsistenceMoist_Hor6", self.Consistence_Moist_CB_Hor_6.currentText())
        self.horizonRecord.setValue("ConsistenceStickiness_Hor6", self.Consistence_Stickiness_CB_Hor_6.currentText())
        self.horizonRecord.setValue("ConsistencePlasticity_Hor6", self.Consistence_Plasticity_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleAbundance1_Hor6", self.Mottle_Abundance_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleSize1_Hor6", self.Mottle_Size_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleContrast1_Hor6", self.Mottle_Contrast_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleShape1_Hor6", self.Mottle_Shape_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleHue1_Hor6", self.Mottle_Hue_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleValue1_Hor6", self.Mottle_Value_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("MottleChroma1_Hor6", self.Mottle_Chroma_1_CB_Hor_6.currentText())
        self.horizonRecord.setValue("RootFine_Hor6", self.Root_Fine_CB_Hor_6.currentText())
        self.horizonRecord.setValue("RootMedium_Hor6", self.Root_Medium_CB_Hor_6.currentText())
        self.horizonRecord.setValue("RootCoarse_Hor6", self.Root_Coarse_CB_Hor_6.currentText())

        # Insert Entry Time and Modified Time
        self.horizonRecord.setValue("Entry_Time", str(QDateTime.currentDateTime().toString("hh:mm:ss dd/MM/yyyy")))

        if self.horizonModel.insertRecord(-1, self.horizonRecord):
            self.horizonModel.submitAll()
        
        self.Insert_PB.setEnabled(False)
        self.Update_PB.setEnabled(True)
        # TODO: check if work
        # Set the horizonMapper to Last index
        self.horizonMapper.toLast()
        # Message box for inserting data
        insertMessageBox = QMessageBox()
        insertMessageBox.setText("Data inserted to Database")
        insertMessageBox.setWindowTitle("Insert Data")
        insertMessageBox.setStandardButtons(QMessageBox.Ok)
        insertMessageBox.exec()

    def updateData(self):
        logging.info("Horizon Windows - Update Button Clicked")

        self.NoFormText = self.Number_Form_CB_Hor.currentText()
        row = self.Number_Form_CB_Hor.findText(self.NoFormText)
        
        self.horizonModel.setData(self.horizonModel.index(row, 1), self.NoFormText)

        ### Horizon 1
        self.horizonModel.setData(self.horizonModel.index(row, 2), self.Hor_Design_Discon_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 3), self.Hor_Design_Master_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 4), self.Hor_Design_Sub_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 5), self.Hor_Design_Number_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 6), self.Hor_Upper_From_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 7), self.Hor_Upper_To_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 8), self.Hor_Lower_From_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 9), self.Hor_Lower_To_LE_Hor_1.text())
        self.horizonModel.setData(self.horizonModel.index(row, 10), self.Boundary_Dist_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 11), self.Boundary_Topo_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 13), self.Soil_Color_Hue_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 14), self.Soil_Color_Value_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 15), self.Soil_Color_Chroma_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 17), self.Soil_Color_Hue_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 18), self.Soil_Color_Value_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 19), self.Soil_Color_Chroma_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 21), self.Soil_Color_Hue_3_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 22), self.Soil_Color_Value_3_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 23), self.Soil_Color_Chroma_3_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 24), self.Texture_Class_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 25), self.Texture_Sand_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 26), self.Texture_Mod_Size_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 27), self.Texture_Mod_Abundance_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 28), self.Structure_Shape_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 29), self.Structure_Size_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 30), self.Structure_Grade_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 31), self.Structure_Relation_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 32), self.Structure_Shape_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 33), self.Structure_Size_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 34), self.Structure_Grade_2_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 35), self.Consistence_Moist_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 36), self.Consistence_Stickiness_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 37), self.Consistence_Plasticity_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 38), self.Mottle_Abundance_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 39), self.Mottle_Size_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 40), self.Mottle_Contrast_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 41), self.Mottle_Shape_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 42), self.Mottle_Hue_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 43), self.Mottle_Value_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 44), self.Mottle_Chroma_1_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 45), self.Root_Fine_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 46), self.Root_Medium_CB_Hor_1.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 47), self.Root_Coarse_CB_Hor_1.currentText())

        ### Horizon 2
        self.horizonModel.setData(self.horizonModel.index(row, 48), self.Hor_Design_Discon_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 49), self.Hor_Design_Master_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 50), self.Hor_Design_Sub_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 51), self.Hor_Design_Number_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 52), self.Hor_Upper_From_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 53), self.Hor_Upper_To_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 54), self.Hor_Lower_From_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 55), self.Hor_Lower_To_LE_Hor_2.text())
        self.horizonModel.setData(self.horizonModel.index(row, 56), self.Boundary_Dist_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 57), self.Boundary_Topo_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 59), self.Soil_Color_Hue_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 60), self.Soil_Color_Value_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 61), self.Soil_Color_Chroma_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 63), self.Soil_Color_Hue_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 64), self.Soil_Color_Value_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 65), self.Soil_Color_Chroma_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 67), self.Soil_Color_Hue_3_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 68), self.Soil_Color_Value_3_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 69), self.Soil_Color_Chroma_3_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 70), self.Texture_Class_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 71), self.Texture_Sand_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 72), self.Texture_Mod_Size_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 73), self.Texture_Mod_Abundance_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 74), self.Structure_Shape_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 75), self.Structure_Size_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 76), self.Structure_Grade_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 77), self.Structure_Relation_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 78), self.Structure_Shape_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 79), self.Structure_Size_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 80), self.Structure_Grade_2_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 81), self.Consistence_Moist_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 82), self.Consistence_Stickiness_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 83), self.Consistence_Plasticity_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 84), self.Mottle_Abundance_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 85), self.Mottle_Size_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 86), self.Mottle_Contrast_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 87), self.Mottle_Shape_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 88), self.Mottle_Hue_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 89), self.Mottle_Value_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 90), self.Mottle_Chroma_1_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 91), self.Root_Fine_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 92), self.Root_Medium_CB_Hor_2.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 93), self.Root_Coarse_CB_Hor_2.currentText())

        ### Horizon 3
        self.horizonModel.setData(self.horizonModel.index(row, 94), self.Hor_Design_Discon_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 95), self.Hor_Design_Master_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 96), self.Hor_Design_Sub_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 97), self.Hor_Design_Number_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 98), self.Hor_Upper_From_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 99), self.Hor_Upper_To_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 100), self.Hor_Lower_From_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 101), self.Hor_Lower_To_LE_Hor_3.text())
        self.horizonModel.setData(self.horizonModel.index(row, 102), self.Boundary_Dist_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 103), self.Boundary_Topo_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 105), self.Soil_Color_Hue_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 106), self.Soil_Color_Value_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 107), self.Soil_Color_Chroma_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 109), self.Soil_Color_Hue_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 110), self.Soil_Color_Value_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 111), self.Soil_Color_Chroma_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 113), self.Soil_Color_Hue_3_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 114), self.Soil_Color_Value_3_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 115), self.Soil_Color_Chroma_3_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 116), self.Texture_Class_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 117), self.Texture_Sand_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 118), self.Texture_Mod_Size_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 119), self.Texture_Mod_Abundance_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 120), self.Structure_Shape_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 121), self.Structure_Size_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 122), self.Structure_Grade_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 123), self.Structure_Relation_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 124), self.Structure_Shape_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 125), self.Structure_Size_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 126), self.Structure_Grade_2_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 127), self.Consistence_Moist_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 128), self.Consistence_Stickiness_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 129), self.Consistence_Plasticity_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 130), self.Mottle_Abundance_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 131), self.Mottle_Size_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 132), self.Mottle_Contrast_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 133), self.Mottle_Shape_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 134), self.Mottle_Hue_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 135), self.Mottle_Value_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 136), self.Mottle_Chroma_1_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 137), self.Root_Fine_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 138), self.Root_Medium_CB_Hor_3.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 139), self.Root_Coarse_CB_Hor_3.currentText())

        ### Horizon 4
        self.horizonModel.setData(self.horizonModel.index(row, 140), self.Hor_Design_Discon_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 141), self.Hor_Design_Master_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 142), self.Hor_Design_Sub_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 143), self.Hor_Design_Number_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 144), self.Hor_Upper_From_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 145), self.Hor_Upper_To_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 146), self.Hor_Lower_From_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 147), self.Hor_Lower_To_LE_Hor_4.text())
        self.horizonModel.setData(self.horizonModel.index(row, 148), self.Boundary_Dist_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 149), self.Boundary_Topo_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 151), self.Soil_Color_Hue_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 152), self.Soil_Color_Value_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 153), self.Soil_Color_Chroma_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 155), self.Soil_Color_Hue_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 156), self.Soil_Color_Value_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 157), self.Soil_Color_Chroma_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 159), self.Soil_Color_Hue_3_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 160), self.Soil_Color_Value_3_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 161), self.Soil_Color_Chroma_3_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 162), self.Texture_Class_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 163), self.Texture_Sand_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 164), self.Texture_Mod_Size_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 165), self.Texture_Mod_Abundance_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 166), self.Structure_Shape_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 167), self.Structure_Size_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 168), self.Structure_Grade_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 169), self.Structure_Relation_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 170), self.Structure_Shape_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 171), self.Structure_Size_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 172), self.Structure_Grade_2_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 173), self.Consistence_Moist_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 174), self.Consistence_Stickiness_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 175), self.Consistence_Plasticity_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 176), self.Mottle_Abundance_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 177), self.Mottle_Size_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 178), self.Mottle_Contrast_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 179), self.Mottle_Shape_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 180), self.Mottle_Hue_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 181), self.Mottle_Value_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 182), self.Mottle_Chroma_1_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 183), self.Root_Fine_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 184), self.Root_Medium_CB_Hor_4.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 185), self.Root_Coarse_CB_Hor_4.currentText())

        ### Horizon 5
        self.horizonModel.setData(self.horizonModel.index(row, 186), self.Hor_Design_Discon_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 187), self.Hor_Design_Master_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 188), self.Hor_Design_Sub_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 189), self.Hor_Design_Number_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 190), self.Hor_Upper_From_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 191), self.Hor_Upper_To_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 192), self.Hor_Lower_From_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 193), self.Hor_Lower_To_LE_Hor_5.text())
        self.horizonModel.setData(self.horizonModel.index(row, 194), self.Boundary_Dist_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 195), self.Boundary_Topo_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 197), self.Soil_Color_Hue_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 198), self.Soil_Color_Value_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 199), self.Soil_Color_Chroma_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 201), self.Soil_Color_Hue_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 202), self.Soil_Color_Value_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 203), self.Soil_Color_Chroma_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 205), self.Soil_Color_Hue_3_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 206), self.Soil_Color_Value_3_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 207), self.Soil_Color_Chroma_3_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 208), self.Texture_Class_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 209), self.Texture_Sand_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 210), self.Texture_Mod_Size_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 211), self.Texture_Mod_Abundance_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 212), self.Structure_Shape_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 213), self.Structure_Size_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 214), self.Structure_Grade_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 215), self.Structure_Relation_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 216), self.Structure_Shape_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 217), self.Structure_Size_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 218), self.Structure_Grade_2_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 219), self.Consistence_Moist_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 220), self.Consistence_Stickiness_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 221), self.Consistence_Plasticity_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 222), self.Mottle_Abundance_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 223), self.Mottle_Size_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 224), self.Mottle_Contrast_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 225), self.Mottle_Shape_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 226), self.Mottle_Hue_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 227), self.Mottle_Value_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 228), self.Mottle_Chroma_1_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 229), self.Root_Fine_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 230), self.Root_Medium_CB_Hor_5.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 231), self.Root_Coarse_CB_Hor_5.currentText())

        ### Horizon 6
        self.horizonModel.setData(self.horizonModel.index(row, 232), self.Hor_Design_Discon_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 233), self.Hor_Design_Master_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 234), self.Hor_Design_Sub_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 235), self.Hor_Design_Number_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 236), self.Hor_Upper_From_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 237), self.Hor_Upper_To_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 238), self.Hor_Lower_From_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 239), self.Hor_Lower_To_LE_Hor_6.text())
        self.horizonModel.setData(self.horizonModel.index(row, 240), self.Boundary_Dist_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 241), self.Boundary_Topo_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 243), self.Soil_Color_Hue_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 244), self.Soil_Color_Value_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 245), self.Soil_Color_Chroma_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 247), self.Soil_Color_Hue_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 248), self.Soil_Color_Value_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 249), self.Soil_Color_Chroma_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 251), self.Soil_Color_Hue_3_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 252), self.Soil_Color_Value_3_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 253), self.Soil_Color_Chroma_3_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 254), self.Texture_Class_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 255), self.Texture_Sand_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 256), self.Texture_Mod_Size_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 257), self.Texture_Mod_Abundance_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 258), self.Structure_Shape_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 259), self.Structure_Size_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 260), self.Structure_Grade_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 261), self.Structure_Relation_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 262), self.Structure_Shape_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 263), self.Structure_Size_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 264), self.Structure_Grade_2_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 265), self.Consistence_Moist_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 266), self.Consistence_Stickiness_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 267), self.Consistence_Plasticity_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 268), self.Mottle_Abundance_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 269), self.Mottle_Size_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 270), self.Mottle_Contrast_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 271), self.Mottle_Shape_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 272), self.Mottle_Hue_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 273), self.Mottle_Value_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 274), self.Mottle_Chroma_1_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 275), self.Root_Fine_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 276), self.Root_Medium_CB_Hor_6.currentText())
        self.horizonModel.setData(self.horizonModel.index(row, 277), self.Root_Coarse_CB_Hor_6.currentText())

        # Insert Entry Modified/Endit Time
        self.horizonModel.setData(self.horizonModel.index(row, 279), QDateTime.currentDateTime().toString("hh:mm:ss dd/MM/yyyy"))

        self.horizonModel.submitAll()

        # set enable / disable the pushButton
        self.Insert_PB.setEnabled(False)
        self.Update_PB.setEnabled(True)

        # self.Number_Form_CB_Hor.setCurrentIndex(row)
        self.horizonMapper.setCurrentIndex(row)
        # Message box of updating data
        updateMessageBox = QMessageBox()
        updateMessageBox.setText("Data updated to Database")
        updateMessageBox.setWindowTitle("Update Data")
        updateMessageBox.setStandardButtons(QMessageBox.Ok)
        updateMessageBox.exec()

    def reloadData(self):
        logging.info("Horizon Windows - Reload Button Clicked")
        horizonFilter = ''
        initializeModel(self.siteModel, "Site", horizonFilter)
        self.populateDataHor1()
        # noFormText = self.Number_Form_CB_Hor.currentText()
        # checkNoFormText = self.Number_Form_CB_Hor.findText(noFormText)

        # proxy = QSortFilterProxyModel()
        # proxy.setSourceModel(self.horizonModel)
        # proxy.setFilterFixedString(noFormText)
        # matchingIndex = proxy.mapToSource(proxy.index(0,0))
        # horizonRow = matchingIndex.row()

        # siteRow = self.Number_Form_CB_Hor.currentIndex()

        # print("horizon row: ", horizonRow)
        # print("site row: ", siteRow)
        # print("is valid: ", matchingIndex.isValid())
        # print("checkNoFormText: ", checkNoFormText)

    def copyData(self):
        print("copy data")
        # get the current text/value of Number Form
        noFormText = self.Form_Copy_CB.currentText()
        
        ### get the proxy of horizonModel so we can filter it
        ### and get the matchingindex from the source model and
        ### ten get the index row of the filtered text/value
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(self.horizonModel)
        proxy.setFilterFixedString(noFormText)

        # set the filter column using secong column or NoForm (index=1) not the 'id' column or first column
        proxy.setFilterKeyColumn(1)
        # match to the source model
        matchingIndex = proxy.mapToSource(proxy.index(0,0))
        # get the index of the row
        horizonRow = matchingIndex.row()
        print("Horizon-Row:", horizonRow)

        ### Horizon 1
        self.Hor_Design_Discon_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 2))))
        self.Hor_Design_Master_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 3))))
        self.Hor_Design_Sub_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 4))))
        self.Hor_Design_Number_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 5))))
        self.Hor_Upper_From_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 6))))
        self.Hor_Upper_To_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 7))))
        self.Hor_Lower_From_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 8))))
        self.Hor_Lower_To_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 9))))
        self.Boundary_Dist_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 10))))
        self.Boundary_Topo_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 11))))
        self.Soil_Color_Hue_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 13))))
        self.Soil_Color_Value_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 14))))
        self.Soil_Color_Chroma_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 15))))
        self.Soil_Color_Hue_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 17))))
        self.Soil_Color_Value_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 18))))
        self.Soil_Color_Chroma_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 19))))
        self.Soil_Color_Hue_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 21))))
        self.Soil_Color_Value_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 22))))
        self.Soil_Color_Chroma_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 23))))
        self.Texture_Class_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 24))))
        self.Texture_Sand_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 25))))
        self.Texture_Mod_Size_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 26))))
        self.Texture_Mod_Abundance_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 27))))
        self.Structure_Shape_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 28))))
        self.Structure_Size_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 29))))
        self.Structure_Grade_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 30))))
        self.Structure_Relation_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 31))))
        self.Structure_Shape_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 32))))
        self.Structure_Size_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 33))))
        self.Structure_Grade_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 34))))
        self.Consistence_Moist_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 35))))
        self.Consistence_Stickiness_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 36))))
        self.Consistence_Plasticity_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 37))))
        self.Mottle_Abundance_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 38))))
        self.Mottle_Size_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 39))))
        self.Mottle_Contrast_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 40))))
        self.Mottle_Shape_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 41))))
        self.Mottle_Hue_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 42))))
        self.Mottle_Value_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 43))))
        self.Mottle_Chroma_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 44))))
        self.Root_Fine_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 45))))
        self.Root_Medium_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 46))))
        self.Root_Coarse_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 47)))) 
        
        ### Horizon 2
        self.Hor_Design_Discon_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 48))))
        self.Hor_Design_Master_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 49))))
        self.Hor_Design_Sub_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 50))))
        self.Hor_Design_Number_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 51))))
        self.Hor_Upper_From_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 52))))
        self.Hor_Upper_To_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 53))))
        self.Hor_Lower_From_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 54))))
        self.Hor_Lower_To_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 55))))
        self.Boundary_Dist_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 56))))
        self.Boundary_Topo_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 57))))
        self.Soil_Color_Hue_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 59))))
        self.Soil_Color_Value_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 60))))
        self.Soil_Color_Chroma_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 61))))
        self.Soil_Color_Hue_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 63))))
        self.Soil_Color_Value_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 64))))
        self.Soil_Color_Chroma_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 65))))
        self.Soil_Color_Hue_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 67))))
        self.Soil_Color_Value_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 68))))
        self.Soil_Color_Chroma_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 69))))
        self.Texture_Class_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 70))))
        self.Texture_Sand_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 71))))
        self.Texture_Mod_Size_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 72))))
        self.Texture_Mod_Abundance_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 73))))
        self.Structure_Shape_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 74))))
        self.Structure_Size_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 75))))
        self.Structure_Grade_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 76))))
        self.Structure_Relation_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 77))))
        self.Structure_Shape_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 78))))
        self.Structure_Size_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 79))))
        self.Structure_Grade_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 80))))
        self.Consistence_Moist_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 81))))
        self.Consistence_Stickiness_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 82))))
        self.Consistence_Plasticity_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 83))))
        self.Mottle_Abundance_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 84))))
        self.Mottle_Size_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 85))))
        self.Mottle_Contrast_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 86))))
        self.Mottle_Shape_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 87))))
        self.Mottle_Hue_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 88))))
        self.Mottle_Value_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 89))))
        self.Mottle_Chroma_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 90))))
        self.Root_Fine_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 91))))
        self.Root_Medium_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 92))))
        self.Root_Coarse_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 93)))) 
        
        ### Horizon 3
        self.Hor_Design_Discon_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 94))))
        self.Hor_Design_Master_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 95))))
        self.Hor_Design_Sub_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 96))))
        self.Hor_Design_Number_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 97))))
        self.Hor_Upper_From_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 98))))
        self.Hor_Upper_To_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 99))))
        self.Hor_Lower_From_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 100))))
        self.Hor_Lower_To_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 101))))
        self.Boundary_Dist_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 102))))
        self.Boundary_Topo_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 103))))
        self.Soil_Color_Hue_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 105))))
        self.Soil_Color_Value_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 106))))
        self.Soil_Color_Chroma_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 107))))
        self.Soil_Color_Hue_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 109))))
        self.Soil_Color_Value_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 110))))
        self.Soil_Color_Chroma_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 111))))
        self.Soil_Color_Hue_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 113))))
        self.Soil_Color_Value_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 114))))
        self.Soil_Color_Chroma_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 115))))
        self.Texture_Class_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 116))))
        self.Texture_Sand_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 117))))
        self.Texture_Mod_Size_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 118))))
        self.Texture_Mod_Abundance_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 119))))
        self.Structure_Shape_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 120))))
        self.Structure_Size_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 121))))
        self.Structure_Grade_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 122))))
        self.Structure_Relation_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 123))))
        self.Structure_Shape_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 124))))
        self.Structure_Size_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 125))))
        self.Structure_Grade_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 126))))
        self.Consistence_Moist_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 127))))
        self.Consistence_Stickiness_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 128))))
        self.Consistence_Plasticity_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 129))))
        self.Mottle_Abundance_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 130))))
        self.Mottle_Size_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 131))))
        self.Mottle_Contrast_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 132))))
        self.Mottle_Shape_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 133))))
        self.Mottle_Hue_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 134))))
        self.Mottle_Value_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 135))))
        self.Mottle_Chroma_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 136))))
        self.Root_Fine_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 137))))
        self.Root_Medium_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 138))))
        self.Root_Coarse_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 139)))) 
        
        ### Horizon 4
        self.Hor_Design_Discon_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 140))))
        self.Hor_Design_Master_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 141))))
        self.Hor_Design_Sub_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 142))))
        self.Hor_Design_Number_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 143))))
        self.Hor_Upper_From_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 144))))
        self.Hor_Upper_To_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 145))))
        self.Hor_Lower_From_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 146))))
        self.Hor_Lower_To_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 147))))
        self.Boundary_Dist_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 148))))
        self.Boundary_Topo_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 149))))
        self.Soil_Color_Hue_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 151))))
        self.Soil_Color_Value_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 152))))
        self.Soil_Color_Chroma_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 153))))
        self.Soil_Color_Hue_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 155))))
        self.Soil_Color_Value_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 156))))
        self.Soil_Color_Chroma_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 157))))
        self.Soil_Color_Hue_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 159))))
        self.Soil_Color_Value_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 160))))
        self.Soil_Color_Chroma_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 161))))
        self.Texture_Class_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 162))))
        self.Texture_Sand_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 163))))
        self.Texture_Mod_Size_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 164))))
        self.Texture_Mod_Abundance_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 165))))
        self.Structure_Shape_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 166))))
        self.Structure_Size_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 167))))
        self.Structure_Grade_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 168))))
        self.Structure_Relation_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 169))))
        self.Structure_Shape_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 170))))
        self.Structure_Size_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 171))))
        self.Structure_Grade_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 172))))
        self.Consistence_Moist_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 173))))
        self.Consistence_Stickiness_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 174))))
        self.Consistence_Plasticity_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 175))))
        self.Mottle_Abundance_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 176))))
        self.Mottle_Size_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 177))))
        self.Mottle_Contrast_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 178))))
        self.Mottle_Shape_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 179))))
        self.Mottle_Hue_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 180))))
        self.Mottle_Value_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 181))))
        self.Mottle_Chroma_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 182))))
        self.Root_Fine_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 183))))
        self.Root_Medium_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 184))))
        self.Root_Coarse_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 185)))) 
        
        ### Horizon 5
        self.Hor_Design_Discon_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 186))))
        self.Hor_Design_Master_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 187))))
        self.Hor_Design_Sub_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 188))))
        self.Hor_Design_Number_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 189))))
        self.Hor_Upper_From_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 190))))
        self.Hor_Upper_To_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 191))))
        self.Hor_Lower_From_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 192))))
        self.Hor_Lower_To_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 193))))
        self.Boundary_Dist_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 194))))
        self.Boundary_Topo_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 195))))
        self.Soil_Color_Hue_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 197))))
        self.Soil_Color_Value_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 198))))
        self.Soil_Color_Chroma_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 199))))
        self.Soil_Color_Hue_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 201))))
        self.Soil_Color_Value_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 202))))
        self.Soil_Color_Chroma_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 203))))
        self.Soil_Color_Hue_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 205))))
        self.Soil_Color_Value_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 206))))
        self.Soil_Color_Chroma_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 207))))
        self.Texture_Class_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 208))))
        self.Texture_Sand_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 209))))
        self.Texture_Mod_Size_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 210))))
        self.Texture_Mod_Abundance_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 211))))
        self.Structure_Shape_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 212))))
        self.Structure_Size_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 213))))
        self.Structure_Grade_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 214))))
        self.Structure_Relation_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 215))))
        self.Structure_Shape_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 216))))
        self.Structure_Size_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 217))))
        self.Structure_Grade_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 218))))
        self.Consistence_Moist_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 219))))
        self.Consistence_Stickiness_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 220))))
        self.Consistence_Plasticity_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 221))))
        self.Mottle_Abundance_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 222))))
        self.Mottle_Size_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 223))))
        self.Mottle_Contrast_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 224))))
        self.Mottle_Shape_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 225))))
        self.Mottle_Hue_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 226))))
        self.Mottle_Value_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 227))))
        self.Mottle_Chroma_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 228))))
        self.Root_Fine_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 229))))
        self.Root_Medium_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 230))))
        self.Root_Coarse_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 231)))) 
        
        ### Horizon 6
        self.Hor_Design_Discon_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 232))))
        self.Hor_Design_Master_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 233))))
        self.Hor_Design_Sub_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 234))))
        self.Hor_Design_Number_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 235))))
        self.Hor_Upper_From_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 236))))
        self.Hor_Upper_To_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 237))))
        self.Hor_Lower_From_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 238))))
        self.Hor_Lower_To_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 239))))
        self.Boundary_Dist_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 240))))
        self.Boundary_Topo_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 241))))
        self.Soil_Color_Hue_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 243))))
        self.Soil_Color_Value_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 244))))
        self.Soil_Color_Chroma_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 245))))
        self.Soil_Color_Hue_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 247))))
        self.Soil_Color_Value_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 248))))
        self.Soil_Color_Chroma_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 249))))
        self.Soil_Color_Hue_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 251))))
        self.Soil_Color_Value_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 252))))
        self.Soil_Color_Chroma_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 253))))
        self.Texture_Class_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 254))))
        self.Texture_Sand_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 255))))
        self.Texture_Mod_Size_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 256))))
        self.Texture_Mod_Abundance_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 257))))
        self.Structure_Shape_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 258))))
        self.Structure_Size_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 259))))
        self.Structure_Grade_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 260))))
        self.Structure_Relation_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 261))))
        self.Structure_Shape_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 262))))
        self.Structure_Size_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 263))))
        self.Structure_Grade_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 264))))
        self.Consistence_Moist_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 265))))
        self.Consistence_Stickiness_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 266))))
        self.Consistence_Plasticity_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 267))))
        self.Mottle_Abundance_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 268))))
        self.Mottle_Size_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 269))))
        self.Mottle_Contrast_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 270))))
        self.Mottle_Shape_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 271))))
        self.Mottle_Hue_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 272))))
        self.Mottle_Value_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 273))))
        self.Mottle_Chroma_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 274))))
        self.Root_Fine_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 275))))
        self.Root_Medium_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 276))))
        self.Root_Coarse_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 277)))) 
        
        # Message box of copying data from another site number form
        currentForm = self.Number_Form_CB_Hor.currentText()
        copiedForm = self.Form_Copy_CB.currentText()
        copyMessageBox = QMessageBox()
        copyMessageBox.setText("Number Form '{0}' copied to Number Form '{1}'".format( copiedForm, currentForm))
        copyMessageBox.setWindowTitle("Copy Data")
        copyMessageBox.setStandardButtons(QMessageBox.Ok)
        copyMessageBox.exec()

    def noFormTextChanged(self):
        # print("focus lost")
        # get the current text/value of Number Form
        noFormText = self.Number_Form_CB_Hor.currentText()
        print("noFromText: ", noFormText)
        # get the current Index of Number Form
        checkNoFormText = self.Number_Form_CB_Hor.findText(noFormText)
        print("Check No Form Text: ", checkNoFormText)

        ### get the proxy of horizonModel so we can filter it
        ### and get the matchingindex from the source model and 
        ### then get the index row of the filtered text/value
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(self.horizonModel)
        proxy.setFilterFixedString(noFormText)

        ### set the filter to the second coloumn(index=1)/NoForm header so that can't confuse with the id
        # that maybe dont have the same number
        # so that we can entry from 'nomor form' bukan 1 atau lebih misal 100
        proxy.setFilterKeyColumn(1)
        # match to the source model
        matchingIndex = proxy.mapToSource(proxy.index(0,0))
        print("Matching Index: ", matchingIndex)
        # get the index of the row
        horizonRow = matchingIndex.row()
        print("Horizon Row: ", horizonRow)

        # using try except to check is Number Form text selected is in 
        # Number Form Model using "checkNoFormText != 1"
        # try:
        
        # check if the Number Form comboBox is not empty
        if (checkNoFormText != -1):
            print("True")
            self.Provinsi_CB_Hor.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 3))))
            self.Kabupaten_CB_Hor.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 4))))
            self.Kecamatan_CB_Hor.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 5))))
            self.Desa_CB_Hor.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 6))))
            self.Date_CB_Hor.setDate(QDate.fromString(str(self.siteModel.data(self.siteModel.index(checkNoFormText,7))), "dd/MM/yyyy"))
            self.Initial_Name_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 8))))
            self.Observation_Number_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 9))))
            self.Kind_Observation_CB_Hor.setCurrentText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 10))))
            self.UTM_Zone_1_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 11))))
            self.UTM_Zone_2_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 12))))
            self.X_East_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 13))))
            self.Y_North_LE_Hor.setText(str(self.siteModel.data(self.siteModel.index(checkNoFormText, 14))))
            

            # setting the info of site in Horizon tabs
            horModel = self.Number_Form_CB_Hor.model()
            initialName = str(horModel.data(horModel.index(checkNoFormText, 8)))
            observationNumber = str(horModel.data(horModel.index(checkNoFormText, 9)))
            spt = str(horModel.data(horModel.index(checkNoFormText, 2)))

            # Horizon 1 Tab
            self.Number_Form_LE_Hor_1.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_1.setText(initialName + observationNumber)
            self.SPT_LE_Hor_1.setText(spt)
            # Horizon 2 Tab
            self.Number_Form_LE_Hor_2.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_2.setText(initialName + observationNumber)
            self.SPT_LE_Hor_2.setText(spt)
            # Horizon 3 Tab
            self.Number_Form_LE_Hor_3.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_3.setText(initialName + observationNumber)
            self.SPT_LE_Hor_3.setText(spt)
            # Horizon 4 Tab
            self.Number_Form_LE_Hor_4.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_4.setText(initialName + observationNumber)
            self.SPT_LE_Hor_4.setText(spt)
            # Horizon 5 Tab
            self.Number_Form_LE_Hor_5.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_5.setText(initialName + observationNumber)
            self.SPT_LE_Hor_5.setText(spt)
            # Horizon 6 Tab
            self.Number_Form_LE_Hor_6.setText(self.Number_Form_CB_Hor.currentText())
            self.Initial_LE_Hor_6.setText(initialName + observationNumber)
            self.SPT_LE_Hor_6.setText(spt)

            # Check if data in horizon model is not empty
            if (horizonRow != -1):
                ### Horizon 1
                self.Hor_Design_Discon_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 2))))
                self.Hor_Design_Master_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 3))))
                self.Hor_Design_Sub_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 4))))
                self.Hor_Design_Number_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 5))))
                self.Hor_Upper_From_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 6))))
                self.Hor_Upper_To_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 7))))
                self.Hor_Lower_From_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 8))))
                self.Hor_Lower_To_LE_Hor_1.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 9))))
                self.Boundary_Dist_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 10))))
                self.Boundary_Topo_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 11))))
                self.Soil_Color_Hue_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 13))))
                self.Soil_Color_Value_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 14))))
                self.Soil_Color_Chroma_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 15))))
                self.Soil_Color_Hue_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 17))))
                self.Soil_Color_Value_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 18))))
                self.Soil_Color_Chroma_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 19))))
                self.Soil_Color_Hue_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 21))))
                self.Soil_Color_Value_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 22))))
                self.Soil_Color_Chroma_3_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 23))))
                self.Texture_Class_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 24))))
                self.Texture_Sand_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 25))))
                self.Texture_Mod_Size_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 26))))
                self.Texture_Mod_Abundance_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 27))))
                self.Structure_Shape_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 28))))
                self.Structure_Size_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 29))))
                self.Structure_Grade_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 30))))
                self.Structure_Relation_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 31))))
                self.Structure_Shape_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 32))))
                self.Structure_Size_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 33))))
                self.Structure_Grade_2_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 34))))
                self.Consistence_Moist_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 35))))
                self.Consistence_Stickiness_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 36))))
                self.Consistence_Plasticity_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 37))))
                self.Mottle_Abundance_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 38))))
                self.Mottle_Size_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 39))))
                self.Mottle_Contrast_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 40))))
                self.Mottle_Shape_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 41))))
                self.Mottle_Hue_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 42))))
                self.Mottle_Value_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 43))))
                self.Mottle_Chroma_1_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 44))))
                self.Root_Fine_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 45))))
                self.Root_Medium_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 46))))
                self.Root_Coarse_CB_Hor_1.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 47)))) 
               
                ### Horizon 2
                self.Hor_Design_Discon_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 48))))
                logging.info(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 48))))
                self.Hor_Design_Master_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 49))))
                self.Hor_Design_Sub_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 50))))
                self.Hor_Design_Number_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 51))))
                self.Hor_Upper_From_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 52))))
                self.Hor_Upper_To_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 53))))
                self.Hor_Lower_From_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 54))))
                self.Hor_Lower_To_LE_Hor_2.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 55))))
                self.Boundary_Dist_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 56))))
                self.Boundary_Topo_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 57))))
                self.Soil_Color_Hue_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 59))))
                self.Soil_Color_Value_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 60))))
                self.Soil_Color_Chroma_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 61))))
                self.Soil_Color_Hue_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 63))))
                self.Soil_Color_Value_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 64))))
                self.Soil_Color_Chroma_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 65))))
                self.Soil_Color_Hue_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 67))))
                self.Soil_Color_Value_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 68))))
                self.Soil_Color_Chroma_3_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 69))))
                self.Texture_Class_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 70))))
                self.Texture_Sand_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 71))))
                self.Texture_Mod_Size_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 72))))
                self.Texture_Mod_Abundance_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 73))))
                self.Structure_Shape_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 74))))
                self.Structure_Size_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 75))))
                self.Structure_Grade_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 76))))
                self.Structure_Relation_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 77))))
                self.Structure_Shape_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 78))))
                self.Structure_Size_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 79))))
                self.Structure_Grade_2_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 80))))
                self.Consistence_Moist_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 81))))
                self.Consistence_Stickiness_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 82))))
                self.Consistence_Plasticity_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 83))))
                self.Mottle_Abundance_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 84))))
                self.Mottle_Size_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 85))))
                self.Mottle_Contrast_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 86))))
                self.Mottle_Shape_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 87))))
                self.Mottle_Hue_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 88))))
                self.Mottle_Value_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 89))))
                self.Mottle_Chroma_1_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 90))))
                self.Root_Fine_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 91))))
                self.Root_Medium_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 92))))
                self.Root_Coarse_CB_Hor_2.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 93)))) 
                
                ### Horizon 3
                self.Hor_Design_Discon_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 94))))
                self.Hor_Design_Master_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 95))))
                self.Hor_Design_Sub_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 96))))
                self.Hor_Design_Number_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 97))))
                self.Hor_Upper_From_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 98))))
                self.Hor_Upper_To_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 99))))
                self.Hor_Lower_From_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 100))))
                self.Hor_Lower_To_LE_Hor_3.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 101))))
                self.Boundary_Dist_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 102))))
                self.Boundary_Topo_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 103))))
                self.Soil_Color_Hue_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 105))))
                self.Soil_Color_Value_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 106))))
                self.Soil_Color_Chroma_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 107))))
                self.Soil_Color_Hue_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 109))))
                self.Soil_Color_Value_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 110))))
                self.Soil_Color_Chroma_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 111))))
                self.Soil_Color_Hue_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 113))))
                self.Soil_Color_Value_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 114))))
                self.Soil_Color_Chroma_3_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 115))))
                self.Texture_Class_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 116))))
                self.Texture_Sand_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 117))))
                self.Texture_Mod_Size_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 118))))
                self.Texture_Mod_Abundance_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 119))))
                self.Structure_Shape_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 120))))
                self.Structure_Size_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 121))))
                self.Structure_Grade_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 122))))
                self.Structure_Relation_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 123))))
                self.Structure_Shape_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 124))))
                self.Structure_Size_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 125))))
                self.Structure_Grade_2_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 126))))
                self.Consistence_Moist_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 127))))
                self.Consistence_Stickiness_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 128))))
                self.Consistence_Plasticity_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 129))))
                self.Mottle_Abundance_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 130))))
                self.Mottle_Size_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 131))))
                self.Mottle_Contrast_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 132))))
                self.Mottle_Shape_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 133))))
                self.Mottle_Hue_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 134))))
                self.Mottle_Value_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 135))))
                self.Mottle_Chroma_1_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 136))))
                self.Root_Fine_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 137))))
                self.Root_Medium_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 138))))
                self.Root_Coarse_CB_Hor_3.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 139)))) 
                
                ### Horizon 4
                self.Hor_Design_Discon_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 140))))
                self.Hor_Design_Master_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 141))))
                self.Hor_Design_Sub_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 142))))
                self.Hor_Design_Number_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 143))))
                self.Hor_Upper_From_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 144))))
                self.Hor_Upper_To_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 145))))
                self.Hor_Lower_From_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 146))))
                self.Hor_Lower_To_LE_Hor_4.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 147))))
                self.Boundary_Dist_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 148))))
                self.Boundary_Topo_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 149))))
                self.Soil_Color_Hue_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 151))))
                self.Soil_Color_Value_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 152))))
                self.Soil_Color_Chroma_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 153))))
                self.Soil_Color_Hue_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 155))))
                self.Soil_Color_Value_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 156))))
                self.Soil_Color_Chroma_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 157))))
                self.Soil_Color_Hue_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 159))))
                self.Soil_Color_Value_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 160))))
                self.Soil_Color_Chroma_3_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 161))))
                self.Texture_Class_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 162))))
                self.Texture_Sand_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 163))))
                self.Texture_Mod_Size_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 164))))
                self.Texture_Mod_Abundance_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 165))))
                self.Structure_Shape_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 166))))
                self.Structure_Size_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 167))))
                self.Structure_Grade_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 168))))
                self.Structure_Relation_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 169))))
                self.Structure_Shape_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 170))))
                self.Structure_Size_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 171))))
                self.Structure_Grade_2_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 172))))
                self.Consistence_Moist_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 173))))
                self.Consistence_Stickiness_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 174))))
                self.Consistence_Plasticity_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 175))))
                self.Mottle_Abundance_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 176))))
                self.Mottle_Size_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 177))))
                self.Mottle_Contrast_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 178))))
                self.Mottle_Shape_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 179))))
                self.Mottle_Hue_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 180))))
                self.Mottle_Value_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 181))))
                self.Mottle_Chroma_1_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 182))))
                self.Root_Fine_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 183))))
                self.Root_Medium_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 184))))
                self.Root_Coarse_CB_Hor_4.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 185)))) 
                
                ### Horizon 5
                self.Hor_Design_Discon_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 186))))
                self.Hor_Design_Master_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 187))))
                self.Hor_Design_Sub_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 188))))
                self.Hor_Design_Number_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 189))))
                self.Hor_Upper_From_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 190))))
                self.Hor_Upper_To_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 191))))
                self.Hor_Lower_From_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 192))))
                self.Hor_Lower_To_LE_Hor_5.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 193))))
                self.Boundary_Dist_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 194))))
                self.Boundary_Topo_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 195))))
                self.Soil_Color_Hue_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 197))))
                self.Soil_Color_Value_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 198))))
                self.Soil_Color_Chroma_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 199))))
                self.Soil_Color_Hue_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 201))))
                self.Soil_Color_Value_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 202))))
                self.Soil_Color_Chroma_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 203))))
                self.Soil_Color_Hue_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 205))))
                self.Soil_Color_Value_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 206))))
                self.Soil_Color_Chroma_3_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 207))))
                self.Texture_Class_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 208))))
                self.Texture_Sand_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 209))))
                self.Texture_Mod_Size_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 210))))
                self.Texture_Mod_Abundance_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 211))))
                self.Structure_Shape_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 212))))
                self.Structure_Size_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 213))))
                self.Structure_Grade_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 214))))
                self.Structure_Relation_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 215))))
                self.Structure_Shape_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 216))))
                self.Structure_Size_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 217))))
                self.Structure_Grade_2_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 218))))
                self.Consistence_Moist_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 219))))
                self.Consistence_Stickiness_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 220))))
                self.Consistence_Plasticity_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 221))))
                self.Mottle_Abundance_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 222))))
                self.Mottle_Size_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 223))))
                self.Mottle_Contrast_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 224))))
                self.Mottle_Shape_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 225))))
                self.Mottle_Hue_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 226))))
                self.Mottle_Value_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 227))))
                self.Mottle_Chroma_1_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 228))))
                self.Root_Fine_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 229))))
                self.Root_Medium_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 230))))
                self.Root_Coarse_CB_Hor_5.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 231)))) 
                
                ### Horizon 6
                self.Hor_Design_Discon_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 232))))
                self.Hor_Design_Master_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 233))))
                self.Hor_Design_Sub_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 234))))
                self.Hor_Design_Number_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 235))))
                self.Hor_Upper_From_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 236))))
                self.Hor_Upper_To_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 237))))
                self.Hor_Lower_From_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 238))))
                self.Hor_Lower_To_LE_Hor_6.setText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 239))))
                self.Boundary_Dist_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 240))))
                self.Boundary_Topo_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 241))))
                self.Soil_Color_Hue_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 243))))
                self.Soil_Color_Value_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 244))))
                self.Soil_Color_Chroma_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 245))))
                self.Soil_Color_Hue_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 247))))
                self.Soil_Color_Value_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 248))))
                self.Soil_Color_Chroma_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 249))))
                self.Soil_Color_Hue_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 251))))
                self.Soil_Color_Value_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 252))))
                self.Soil_Color_Chroma_3_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 253))))
                self.Texture_Class_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 254))))
                self.Texture_Sand_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 255))))
                self.Texture_Mod_Size_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 256))))
                self.Texture_Mod_Abundance_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 257))))
                self.Structure_Shape_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 258))))
                self.Structure_Size_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 259))))
                self.Structure_Grade_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 260))))
                self.Structure_Relation_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 261))))
                self.Structure_Shape_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 262))))
                self.Structure_Size_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 263))))
                self.Structure_Grade_2_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 264))))
                self.Consistence_Moist_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 265))))
                self.Consistence_Stickiness_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 266))))
                self.Consistence_Plasticity_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 267))))
                self.Mottle_Abundance_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 268))))
                self.Mottle_Size_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 269))))
                self.Mottle_Contrast_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 270))))
                self.Mottle_Shape_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 271))))
                self.Mottle_Hue_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 272))))
                self.Mottle_Value_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 273))))
                self.Mottle_Chroma_1_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 274))))
                self.Root_Fine_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 275))))
                self.Root_Medium_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 276))))
                self.Root_Coarse_CB_Hor_6.setCurrentText(str(self.horizonModel.data(self.horizonModel.index(horizonRow, 277)))) 

                # setting the info of site in Horizon tabs
                # horModel = self.Number_Form_CB_Hor.model()
                # initialName = str(horModel.data(horModel.index(checkNoFormText, 8)))
                # observationNumber = str(horModel.data(horModel.index(checkNoFormText, 9)))
                # self.Number_Form_LE_Hor_1.setText(self.Number_Form_CB_Hor.currentText())
                # self.Initial_LE_Hor_1.setText(initialName + observationNumber)
                # enabled/disabled the insert or/and update push Button
                self.Insert_PB.setEnabled(False)
                self.Update_PB.setEnabled(True)
                # set the copy Button and Combobox to False
                self.Copy_All_PB.setEnabled(False)
                self.Form_Copy_CB.setEnabled(False)

            else:
                # set the copy Button and Combobox to True
                self.Copy_All_PB.setEnabled(True)
                self.Form_Copy_CB.setEnabled(True)
                # set the value of copy comboBox to the last number form index in the SiteModel/HorizonModel
                self.Form_Copy_CB.setCurrentIndex(self.proxyHorizonModel.rowCount()-1)

                # Horizon 1
                self.Hor_Design_Discon_LE_Hor_1.setText("")
                self.Hor_Design_Master_LE_Hor_1.setText("")
                self.Hor_Design_Sub_LE_Hor_1.setText("")
                self.Hor_Design_Number_LE_Hor_1.setText("")
                self.Hor_Upper_From_LE_Hor_1.setText("")
                self.Hor_Upper_To_LE_Hor_1.setText("")
                self.Hor_Lower_From_LE_Hor_1.setText("")
                self.Hor_Lower_To_LE_Hor_1.setText("")
                self.Boundary_Dist_CB_Hor_1.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_1.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_1.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_1.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_1.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_1.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_1.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_1.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_1.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_1.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_1.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_1.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_1.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_1.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_1.setCurrentIndex(-1)

                # Horizon 2
                self.Hor_Design_Discon_LE_Hor_2.setText("")
                self.Hor_Design_Master_LE_Hor_2.setText("")
                self.Hor_Design_Sub_LE_Hor_2.setText("")
                self.Hor_Design_Number_LE_Hor_2.setText("")
                self.Hor_Upper_From_LE_Hor_2.setText("")
                self.Hor_Upper_To_LE_Hor_2.setText("")
                self.Hor_Lower_From_LE_Hor_2.setText("")
                self.Hor_Lower_To_LE_Hor_2.setText("")
                self.Boundary_Dist_CB_Hor_2.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_2.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_2.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_2.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_2.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_2.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_2.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_2.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_2.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_2.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_2.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_2.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_2.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_2.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_2.setCurrentIndex(-1)

                # Horizon 3
                self.Hor_Design_Discon_LE_Hor_3.setText("")
                self.Hor_Design_Master_LE_Hor_3.setText("")
                self.Hor_Design_Sub_LE_Hor_3.setText("")
                self.Hor_Design_Number_LE_Hor_3.setText("")
                self.Hor_Upper_From_LE_Hor_3.setText("")
                self.Hor_Upper_To_LE_Hor_3.setText("")
                self.Hor_Lower_From_LE_Hor_3.setText("")
                self.Hor_Lower_To_LE_Hor_3.setText("")
                self.Boundary_Dist_CB_Hor_3.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_3.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_3.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_3.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_3.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_3.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_3.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_3.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_3.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_3.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_3.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_3.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_3.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_3.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_3.setCurrentIndex(-1)

                # Horizon 4
                self.Hor_Design_Discon_LE_Hor_4.setText("")
                self.Hor_Design_Master_LE_Hor_4.setText("")
                self.Hor_Design_Sub_LE_Hor_4.setText("")
                self.Hor_Design_Number_LE_Hor_4.setText("")
                self.Hor_Upper_From_LE_Hor_4.setText("")
                self.Hor_Upper_To_LE_Hor_4.setText("")
                self.Hor_Lower_From_LE_Hor_4.setText("")
                self.Hor_Lower_To_LE_Hor_4.setText("")
                self.Boundary_Dist_CB_Hor_4.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_4.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_4.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_4.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_4.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_4.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_4.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_4.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_4.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_4.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_4.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_4.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_4.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_4.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_4.setCurrentIndex(-1)

                # Horizon 5
                self.Hor_Design_Discon_LE_Hor_5.setText("")
                self.Hor_Design_Master_LE_Hor_5.setText("")
                self.Hor_Design_Sub_LE_Hor_5.setText("")
                self.Hor_Design_Number_LE_Hor_5.setText("")
                self.Hor_Upper_From_LE_Hor_5.setText("")
                self.Hor_Upper_To_LE_Hor_5.setText("")
                self.Hor_Lower_From_LE_Hor_5.setText("")
                self.Hor_Lower_To_LE_Hor_5.setText("")
                self.Boundary_Dist_CB_Hor_5.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_5.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_5.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_5.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_5.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_5.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_5.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_5.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_5.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_5.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_5.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_5.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_5.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_5.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_5.setCurrentIndex(-1)
                
                # Horizon 6
                self.Hor_Design_Discon_LE_Hor_6.setText("")
                self.Hor_Design_Master_LE_Hor_6.setText("")
                self.Hor_Design_Sub_LE_Hor_6.setText("")
                self.Hor_Design_Number_LE_Hor_6.setText("")
                self.Hor_Upper_From_LE_Hor_6.setText("")
                self.Hor_Upper_To_LE_Hor_6.setText("")
                self.Hor_Lower_From_LE_Hor_6.setText("")
                self.Hor_Lower_To_LE_Hor_6.setText("")
                self.Boundary_Dist_CB_Hor_6.setCurrentIndex(-1)
                self.Boundary_Topo_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Hue_1_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Value_1_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Chroma_1_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Hue_2_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Value_2_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Chroma_2_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Hue_3_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Value_3_CB_Hor_6.setCurrentIndex(-1)
                self.Soil_Color_Chroma_3_CB_Hor_6.setCurrentIndex(-1)
                self.Texture_Class_CB_Hor_6.setCurrentIndex(-1)
                self.Texture_Sand_CB_Hor_6.setCurrentIndex(-1)
                self.Texture_Mod_Size_CB_Hor_6.setCurrentIndex(-1)
                self.Texture_Mod_Abundance_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Shape_1_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Size_1_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Grade_1_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Relation_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Shape_2_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Size_2_CB_Hor_6.setCurrentIndex(-1)
                self.Structure_Grade_2_CB_Hor_6.setCurrentIndex(-1)
                self.Consistence_Moist_CB_Hor_6.setCurrentIndex(-1)
                self.Consistence_Stickiness_CB_Hor_6.setCurrentIndex(-1)
                self.Consistence_Plasticity_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Abundance_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Size_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Contrast_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Shape_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Hue_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Value_1_CB_Hor_6.setCurrentIndex(-1)
                self.Mottle_Chroma_1_CB_Hor_6.setCurrentIndex(-1)
                self.Root_Fine_CB_Hor_6.setCurrentIndex(-1)
                self.Root_Medium_CB_Hor_6.setCurrentIndex(-1)
                self.Root_Coarse_CB_Hor_6.setCurrentIndex(-1)
                
                self.Insert_PB.setEnabled(True)
                self.Update_PB.setEnabled(False)                     

            # self.Insert_PB.setEnabled(False)
            # self.Update_PB.setEnabled(True)
            self.siteMapper.setCurrentIndex(checkNoFormText)
            # self.horizonMapper.setCurrentIndex(checkNoFormText)
            # TODO: check siteMapper with horizonMapper is not working correctlly

            #setting Next_PB dan Previous_PB setEnabled(True)
            self.Next_PB.setEnabled(True)
            self.Previous_PB.setEnabled(True)

            

        else:
            print("Index horizon FALSE")
            # set the combobox or lineedit to empty or blank Site Tab
            self.Provinsi_CB_Hor.setCurrentIndex(-1)
            self.Kabupaten_CB_Hor.setCurrentIndex(-1)
            self.Kecamatan_CB_Hor.setCurrentIndex(-1)
            self.Desa_CB_Hor.setCurrentIndex(-1)
            self.Date_CB_Hor.setDate(QDate.fromString(str(self.siteModel.data(self.siteModel.index(checkNoFormText,7))), "dd/MM/yyyy"))
            self.Initial_Name_LE_Hor.setText("")
            self.Observation_Number_LE_Hor.setText("")
            self.Kind_Observation_CB_Hor.setCurrentIndex(-1)
            self.UTM_Zone_1_LE_Hor.setText("")
            self.UTM_Zone_2_LE_Hor.setText("")
            self.X_East_LE_Hor.setText("")
            self.Y_North_LE_Hor.setText("")


            ### Set Horizon 1 Field to Empty
            self.Hor_Design_Discon_LE_Hor_1.setText("")
            self.Hor_Design_Master_LE_Hor_1.setText("")
            self.Hor_Design_Sub_LE_Hor_1.setText("")
            self.Hor_Design_Number_LE_Hor_1.setText("")
            self.Hor_Upper_From_LE_Hor_1.setText("")
            self.Hor_Upper_To_LE_Hor_1.setText("")
            self.Hor_Lower_From_LE_Hor_1.setText("")
            self.Hor_Lower_To_LE_Hor_1.setText("")
            self.Boundary_Dist_CB_Hor_1.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_1.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_1.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_1.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_1.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_1.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_1.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_1.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_1.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_1.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_1.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_1.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_1.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_1.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_1.setCurrentIndex(-1) 

            # Horizon 2
            self.Hor_Design_Discon_LE_Hor_2.setText("")
            self.Hor_Design_Master_LE_Hor_2.setText("")
            self.Hor_Design_Sub_LE_Hor_2.setText("")
            self.Hor_Design_Number_LE_Hor_2.setText("")
            self.Hor_Upper_From_LE_Hor_2.setText("")
            self.Hor_Upper_To_LE_Hor_2.setText("")
            self.Hor_Lower_From_LE_Hor_2.setText("")
            self.Hor_Lower_To_LE_Hor_2.setText("")
            self.Boundary_Dist_CB_Hor_2.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_2.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_2.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_2.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_2.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_2.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_2.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_2.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_2.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_2.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_2.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_2.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_2.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_2.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_2.setCurrentIndex(-1)

            # Horizon 3
            self.Hor_Design_Discon_LE_Hor_3.setText("")
            self.Hor_Design_Master_LE_Hor_3.setText("")
            self.Hor_Design_Sub_LE_Hor_3.setText("")
            self.Hor_Design_Number_LE_Hor_3.setText("")
            self.Hor_Upper_From_LE_Hor_3.setText("")
            self.Hor_Upper_To_LE_Hor_3.setText("")
            self.Hor_Lower_From_LE_Hor_3.setText("")
            self.Hor_Lower_To_LE_Hor_3.setText("")
            self.Boundary_Dist_CB_Hor_3.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_3.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_3.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_3.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_3.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_3.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_3.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_3.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_3.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_3.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_3.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_3.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_3.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_3.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_3.setCurrentIndex(-1)

            # Horizon 4
            self.Hor_Design_Discon_LE_Hor_4.setText("")
            self.Hor_Design_Master_LE_Hor_4.setText("")
            self.Hor_Design_Sub_LE_Hor_4.setText("")
            self.Hor_Design_Number_LE_Hor_4.setText("")
            self.Hor_Upper_From_LE_Hor_4.setText("")
            self.Hor_Upper_To_LE_Hor_4.setText("")
            self.Hor_Lower_From_LE_Hor_4.setText("")
            self.Hor_Lower_To_LE_Hor_4.setText("")
            self.Boundary_Dist_CB_Hor_4.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_4.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_4.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_4.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_4.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_4.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_4.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_4.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_4.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_4.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_4.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_4.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_4.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_4.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_4.setCurrentIndex(-1)

            # Horizon 5
            self.Hor_Design_Discon_LE_Hor_5.setText("")
            self.Hor_Design_Master_LE_Hor_5.setText("")
            self.Hor_Design_Sub_LE_Hor_5.setText("")
            self.Hor_Design_Number_LE_Hor_5.setText("")
            self.Hor_Upper_From_LE_Hor_5.setText("")
            self.Hor_Upper_To_LE_Hor_5.setText("")
            self.Hor_Lower_From_LE_Hor_5.setText("")
            self.Hor_Lower_To_LE_Hor_5.setText("")
            self.Boundary_Dist_CB_Hor_5.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_5.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_5.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_5.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_5.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_5.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_5.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_5.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_5.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_5.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_5.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_5.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_5.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_5.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_5.setCurrentIndex(-1)
            
            # Horizon 6
            self.Hor_Design_Discon_LE_Hor_6.setText("")
            self.Hor_Design_Master_LE_Hor_6.setText("")
            self.Hor_Design_Sub_LE_Hor_6.setText("")
            self.Hor_Design_Number_LE_Hor_6.setText("")
            self.Hor_Upper_From_LE_Hor_6.setText("")
            self.Hor_Upper_To_LE_Hor_6.setText("")
            self.Hor_Lower_From_LE_Hor_6.setText("")
            self.Hor_Lower_To_LE_Hor_6.setText("")
            self.Boundary_Dist_CB_Hor_6.setCurrentIndex(-1)
            self.Boundary_Topo_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Hue_1_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Value_1_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Chroma_1_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Hue_2_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Value_2_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Chroma_2_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Hue_3_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Value_3_CB_Hor_6.setCurrentIndex(-1)
            self.Soil_Color_Chroma_3_CB_Hor_6.setCurrentIndex(-1)
            self.Texture_Class_CB_Hor_6.setCurrentIndex(-1)
            self.Texture_Sand_CB_Hor_6.setCurrentIndex(-1)
            self.Texture_Mod_Size_CB_Hor_6.setCurrentIndex(-1)
            self.Texture_Mod_Abundance_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Shape_1_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Size_1_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Grade_1_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Relation_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Shape_2_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Size_2_CB_Hor_6.setCurrentIndex(-1)
            self.Structure_Grade_2_CB_Hor_6.setCurrentIndex(-1)
            self.Consistence_Moist_CB_Hor_6.setCurrentIndex(-1)
            self.Consistence_Stickiness_CB_Hor_6.setCurrentIndex(-1)
            self.Consistence_Plasticity_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Abundance_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Size_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Contrast_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Shape_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Hue_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Value_1_CB_Hor_6.setCurrentIndex(-1)
            self.Mottle_Chroma_1_CB_Hor_6.setCurrentIndex(-1)
            self.Root_Fine_CB_Hor_6.setCurrentIndex(-1)
            self.Root_Medium_CB_Hor_6.setCurrentIndex(-1)
            self.Root_Coarse_CB_Hor_6.setCurrentIndex(-1)
                
            self.Insert_PB.setEnabled(True)
            self.Update_PB.setEnabled(False)
        # except:

        print("error")

        # print(checkNoFormText)

    def setEvent(self):
        ''' EVENT '''
        logging.info("Horizon Windows - Set Events")
        # Copy form another number form events
        self.Copy_All_PB.clicked.connect(self.copyData)

        # Current text changed
        self.Number_Form_CB_Hor.currentTextChanged.connect(self.noFormTextChanged)

        ### FOCUS OUT
        ### Install EventFilter
        self.focusOutHorUpperFrom_Hor1 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_1.installEventFilter(self.focusOutHorUpperFrom_Hor1)
        self.focusOutHorLowerFrom_Hor1 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_1.installEventFilter(self.focusOutHorLowerFrom_Hor1)
        self.focusOutHorUpperFrom_Hor2 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_2.installEventFilter(self.focusOutHorUpperFrom_Hor2)
        self.focusOutHorLowerFrom_Hor2 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_2.installEventFilter(self.focusOutHorLowerFrom_Hor2)
        self.focusOutHorUpperFrom_Hor3 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_3.installEventFilter(self.focusOutHorUpperFrom_Hor3)
        self.focusOutHorLowerFrom_Hor3 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_3.installEventFilter(self.focusOutHorLowerFrom_Hor3)
        self.focusOutHorUpperFrom_Hor4 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_4.installEventFilter(self.focusOutHorUpperFrom_Hor4)
        self.focusOutHorLowerFrom_Hor4 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_4.installEventFilter(self.focusOutHorLowerFrom_Hor4)
        self.focusOutHorUpperFrom_Hor5 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_5.installEventFilter(self.focusOutHorUpperFrom_Hor5)
        self.focusOutHorLowerFrom_Hor5 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_5.installEventFilter(self.focusOutHorLowerFrom_Hor5)
        self.focusOutHorUpperFrom_Hor6 = FocusOutFilter()
        self.Hor_Upper_From_LE_Hor_6.installEventFilter(self.focusOutHorUpperFrom_Hor6)
        self.focusOutHorLowerFrom_Hor6 = FocusOutFilter()
        self.Hor_Lower_From_LE_Hor_6.installEventFilter(self.focusOutHorLowerFrom_Hor6)

        ### Focus out event
        self.focusOutHorUpperFrom_Hor1.focusOut.connect(self.horUpperFromFocusOut_Hor1)
        self.focusOutHorLowerFrom_Hor1.focusOut.connect(self.horLowerFromFocusOut_Hor1)
        self.focusOutHorUpperFrom_Hor2.focusOut.connect(self.horUpperFromFocusOut_Hor2)
        self.focusOutHorLowerFrom_Hor2.focusOut.connect(self.horLowerFromFocusOut_Hor2)
        self.focusOutHorUpperFrom_Hor3.focusOut.connect(self.horUpperFromFocusOut_Hor3)
        self.focusOutHorLowerFrom_Hor3.focusOut.connect(self.horLowerFromFocusOut_Hor3)
        self.focusOutHorUpperFrom_Hor4.focusOut.connect(self.horUpperFromFocusOut_Hor4)
        self.focusOutHorLowerFrom_Hor4.focusOut.connect(self.horLowerFromFocusOut_Hor4)
        self.focusOutHorUpperFrom_Hor5.focusOut.connect(self.horUpperFromFocusOut_Hor5)
        self.focusOutHorLowerFrom_Hor5.focusOut.connect(self.horLowerFromFocusOut_Hor5)
        self.focusOutHorUpperFrom_Hor6.focusOut.connect(self.horUpperFromFocusOut_Hor6)
        self.focusOutHorLowerFrom_Hor6.focusOut.connect(self.horLowerFromFocusOut_Hor6)

        # SOIL COLOR HORIZON 1
        self.Soil_Color_Hue_1_CB_Hor_1.currentTextChanged.connect(self.Hue_1_Hor_1_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_1.currentTextChanged.connect(self.Hue_2_Hor_1_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_1.currentTextChanged.connect(self.Hue_3_Hor_1_TextChanged)
        self.Mottle_Hue_1_CB_Hor_1.currentTextChanged.connect(self.mottleHueTextChanged_Hor1)
        # SOIL COLOR HORIZON 2
        self.Soil_Color_Hue_1_CB_Hor_2.currentTextChanged.connect(self.Hue_1_Hor_2_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_2.currentTextChanged.connect(self.Hue_2_Hor_2_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_2.currentTextChanged.connect(self.Hue_3_Hor_2_TextChanged)
        self.Mottle_Hue_1_CB_Hor_2.currentTextChanged.connect(self.mottleHueTextChanged_Hor2)
        # SOIL COLOR HORIZON 3
        self.Soil_Color_Hue_1_CB_Hor_3.currentTextChanged.connect(self.Hue_1_Hor_3_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_3.currentTextChanged.connect(self.Hue_2_Hor_3_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_3.currentTextChanged.connect(self.Hue_3_Hor_3_TextChanged)
        self.Mottle_Hue_1_CB_Hor_3.currentTextChanged.connect(self.mottleHueTextChanged_Hor3)
        # SOIL COLOR HORIZON 4
        self.Soil_Color_Hue_1_CB_Hor_4.currentTextChanged.connect(self.Hue_1_Hor_4_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_4.currentTextChanged.connect(self.Hue_2_Hor_4_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_4.currentTextChanged.connect(self.Hue_3_Hor_4_TextChanged)
        self.Mottle_Hue_1_CB_Hor_4.currentTextChanged.connect(self.mottleHueTextChanged_Hor4)
        # SOIL COLOR HORIZON 5
        self.Soil_Color_Hue_1_CB_Hor_5.currentTextChanged.connect(self.Hue_1_Hor_5_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_5.currentTextChanged.connect(self.Hue_2_Hor_5_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_5.currentTextChanged.connect(self.Hue_3_Hor_5_TextChanged)
        self.Mottle_Hue_1_CB_Hor_5.currentTextChanged.connect(self.mottleHueTextChanged_Hor5)
        # SOIL COLOR HORIZON 6
        self.Soil_Color_Hue_1_CB_Hor_6.currentTextChanged.connect(self.Hue_1_Hor_6_TextChanged)
        self.Soil_Color_Hue_2_CB_Hor_6.currentTextChanged.connect(self.Hue_2_Hor_6_TextChanged)
        self.Soil_Color_Hue_3_CB_Hor_6.currentTextChanged.connect(self.Hue_3_Hor_6_TextChanged)
        self.Mottle_Hue_1_CB_Hor_6.currentTextChanged.connect(self.mottleHueTextChanged_Hor6)

        # self.Number_Form_CB_Hor.activated.connect(self.noFormActivated)
        self.Insert_PB.clicked.connect(self.insertData)
        self.Update_PB.clicked.connect(self.updateData)
        self.Reload_PB.clicked.connect(self.reloadData)
        self.Delete_PB.clicked.connect(self.deleteRowData)
        self.First_PB.clicked.connect(self.toFirst)
        self.Previous_PB.clicked.connect(self.toPrevious)
        self.Next_PB.clicked.connect(self.toNext)
        self.Last_PB.clicked.connect(self.toLast)
        self.ShowTable_PB.clicked.connect(self.showTable)

        # self.To_Horizon_1_PB.clicked.connect(self.toHorizon1)
        # self.To_Horizon_2_PB.clicked.connect(self.toHorizon2)
        # self.To_Site_Info_PB.clicked.connect(self.toSiteInfo)

        # self.Hor_Upper_From_LE_Hor_1.focusOut
    
    def showTable(self):

        logging.info("Horizon Windows - Show database tables")

        print("showTable")
        dialog = QDialog()
        vLayout = QVBoxLayout()
        tableView = QTableView()
        tableView.setModel(self.horizonModel)
        # label = QLabel("Test")
        # pB = QPushButton("Test Clicked")
        vLayout.addWidget(tableView)
        # vLayout.addWidget(label)
        # vLayout.addWidget(pB)
        dialog.setLayout(vLayout)
        dialog.setWindowTitle("Horizon Tables")
        dialog.setGeometry(130, 65, 1100, 600)
        dialog.exec()
    
    def toFirst(self):
        logging.info("Horizon Windows - Move to the first index")

        self.horizonMapper.toFirst()

        ## set the GroupBox Site Information in horizon 1 Tab
        self.Number_Form_LE_Hor_1.setText(str(self.Number_Form_CB_Hor.currentText()))
        self.Initial_LE_Hor_1.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 2 Tab
        self.Number_Form_LE_Hor_2.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_2.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 3 Tab
        self.Number_Form_LE_Hor_3.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_3.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 4 Tab
        self.Number_Form_LE_Hor_4.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_4.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 5 Tab
        self.Number_Form_LE_Hor_5.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_5.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 6 Tab
        self.Number_Form_LE_Hor_6.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_6.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        
        ## setting Next_PB dan Previous_PB setEnabled(True)
        self.Next_PB.setEnabled(True)
        self.Previous_PB.setEnabled(True)
    
    def toPrevious(self):
        logging.info("Horizon Windows - Move to the previous index")

        # TODO: if NoForm more than horizonModel length/count
        # then NoForm index is the same with horizonModel length
        self.horizonMapper.toPrevious()
        # self.siteMapper.toPrevious()
       
        ## set the GroupBox Site Information in horizon 1 Tab
        self.Number_Form_LE_Hor_1.setText(str(self.Number_Form_CB_Hor.currentText()))
        self.Initial_LE_Hor_1.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 2 Tab
        self.Number_Form_LE_Hor_2.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_2.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 3 Tab
        self.Number_Form_LE_Hor_3.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_3.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 4 Tab
        self.Number_Form_LE_Hor_4.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_4.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 5 Tab
        self.Number_Form_LE_Hor_5.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_5.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 6 Tab
        self.Number_Form_LE_Hor_6.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_6.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        
    def toNext(self):
        # To limit siteMapper.toNext() if the number form in horizonModel/mapper
        logging.info("Horizon Windows - Move to the Next index")

        # is less than siteMapper/Model

        noFormText = self.Number_Form_CB_Hor.currentText()
        intNoFormText = int(noFormText)
        
        ### Dont need if statement if the first number form start not from 1
        ### toNext function is not working. Because intNoFormText always > horizon row count.
        # if ( intNoFormText < self.horizonModel.rowCount()):
        #     print("True")
        self.horizonMapper.toNext()
            # self.siteMapper.toNext()

        ## set the GroupBox Site Information in 
        # Horizon 1 Tab
        self.Number_Form_LE_Hor_1.setText(str(self.Number_Form_CB_Hor.currentText()))
        self.Initial_LE_Hor_1.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 2 Tab
        self.Number_Form_LE_Hor_2.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_2.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 3 Tab
        self.Number_Form_LE_Hor_3.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_3.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 4 Tab
        self.Number_Form_LE_Hor_4.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_4.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 5 Tab
        self.Number_Form_LE_Hor_5.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_5.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 6 Tab
        self.Number_Form_LE_Hor_6.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_6.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        
    def toLast(self):
        logging.info("Horizon Windows - Move to the Last index")

        self.horizonMapper.toLast()
        # self.siteMapper.toLast()
        
        ## set the GroupBox Site Information in horizon 1 Tab
        self.Number_Form_LE_Hor_1.setText(str(self.Number_Form_CB_Hor.currentText()))
        self.Initial_LE_Hor_1.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 2 Tab
        self.Number_Form_LE_Hor_2.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_2.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 3 Tab
        self.Number_Form_LE_Hor_3.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_3.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 4 Tab
        self.Number_Form_LE_Hor_4.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_4.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 5 Tab
        self.Number_Form_LE_Hor_5.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_5.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        # Horizon 6 Tab
        self.Number_Form_LE_Hor_6.setText(self.Number_Form_CB_Hor.currentText())
        self.Initial_LE_Hor_6.setText(self.Initial_Name_LE_Hor.text() + self.Observation_Number_LE_Hor.text())
        
        #setting Next_PB dan Previous_PB setEnabled(True)
        self.Next_PB.setEnabled(True)
        self.Previous_PB.setEnabled(True)
        
    def deleteRowData(self):
        logging.info("Horizon Windows - Delete row data button clicked")

        ## Untuk menentukan baris/row berdasarkan number form (TAPI SALAH GUNAKAN CODE DIBAWAH)
        # numberFormValue = self.Number_Form_CB_Hor.currentText()
        # row = self.Number_Form_CB_Hor.findText(numberFormValue)

        noFormText = self.Number_Form_CB_Hor.currentText()

        ### get the proxy of horizonModel so we can filter it
        ### and get the matchingindex from the source model and 
        ### then get the index row of the filtered text/value
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(self.horizonModel)
        proxy.setFilterFixedString(noFormText)

        ### set the filter to the second coloumn(index=1)/NoForm header so that can't confuse with the id
        # that maybe dont have the same number
        # so that we can entry from 'nomor form' bukan 1 atau lebih misal 100
        proxy.setFilterKeyColumn(1)
        # match to the source model
        matchingIndex = proxy.mapToSource(proxy.index(0,0))
        print("Matching Index: ", matchingIndex)
        # get the index of the row
        horizonRow = matchingIndex.row()
        print("Horizon Row: ", horizonRow)

        self.horizonModel.removeRow(horizonRow)
        self.horizonModel.submitAll()

        # self.mapper.toPrevious()
        
        # after delete/remove row go to the last index through mapper
        # TODO : I think this is need to check the activated/focusLost event in the Number_Form_CB_Hor
        #        to get the index of the mapper
        
        ### Set the Number Form to the -1 from the deleted row/NoForm
        # previousNoForm = int(check) - 1
        # self.ui.Number_Form_CB_Hor.setCurrentText(str(previousNoForm))


    ### HORIZON 1 SOIL COLOR
    def Hue_1_Hor_1_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_1.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_1.model().clear()
        self.Soil_Color_Value_1_CB_Hor_1.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_1.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_1.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_1, self.Soil_Color_Chroma_1_CB_Hor_1, 
                        self.SoilColorValue1Completer_Hor1, self.SoilColorChroma1Completer_Hor1)
    
    def Hue_2_Hor_1_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_1.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_1.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_1.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_1.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_1.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_1, self.Soil_Color_Chroma_2_CB_Hor_1, 
                        self.SoilColorValue2Completer_Hor1, self.SoilColorChroma2Completer_Hor1)

    def Hue_3_Hor_1_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_1.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_1.model().clear()
        self.Soil_Color_Value_3_CB_Hor_1.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_1.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_1.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_1, self.Soil_Color_Chroma_3_CB_Hor_1, 
                        self.SoilColorValue3Completer_Hor1, self.SoilColorChroma3Completer_Hor1)

    ### HORIZON 2 SOIL COLOR
    def Hue_1_Hor_2_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_2.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_2.model().clear()
        self.Soil_Color_Value_1_CB_Hor_2.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_2.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_2.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_2, self.Soil_Color_Chroma_1_CB_Hor_2, 
                        self.SoilColorValue1Completer_Hor2, self.SoilColorChroma1Completer_Hor2)
    
    def Hue_2_Hor_2_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_2.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_2.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_2.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_2.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_2.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_2, self.Soil_Color_Chroma_2_CB_Hor_2, 
                        self.SoilColorValue2Completer_Hor2, self.SoilColorChroma2Completer_Hor2)

    def Hue_3_Hor_2_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_2.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_2.model().clear()
        self.Soil_Color_Value_3_CB_Hor_2.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_2.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_2.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_2, self.Soil_Color_Chroma_3_CB_Hor_2, 
                        self.SoilColorValue3Completer_Hor2, self.SoilColorChroma3Completer_Hor2)

    ### HORIZON 3 SOIL COLOR
    def Hue_1_Hor_3_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_3.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_3.model().clear()
        self.Soil_Color_Value_1_CB_Hor_3.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_3.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_3.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_3, self.Soil_Color_Chroma_1_CB_Hor_3, 
                        self.SoilColorValue1Completer_Hor3, self.SoilColorChroma1Completer_Hor3)
    
    def Hue_2_Hor_3_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_3.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_3.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_3.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_3.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_3.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_3, self.Soil_Color_Chroma_2_CB_Hor_3, 
                        self.SoilColorValue2Completer_Hor3, self.SoilColorChroma2Completer_Hor3)

    def Hue_3_Hor_3_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_3.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_3.model().clear()
        self.Soil_Color_Value_3_CB_Hor_3.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_3.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_3.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_3, self.Soil_Color_Chroma_3_CB_Hor_3, 
                        self.SoilColorValue3Completer_Hor3, self.SoilColorChroma3Completer_Hor3)

    ### HORIZON 4 SOIL COLOR
    def Hue_1_Hor_4_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_4.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_4.model().clear()
        self.Soil_Color_Value_1_CB_Hor_4.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_4.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_4.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_4, self.Soil_Color_Chroma_1_CB_Hor_4, 
                        self.SoilColorValue1Completer_Hor4, self.SoilColorChroma1Completer_Hor4)
    
    def Hue_2_Hor_4_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_4.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_4.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_4.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_4.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_4.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_4, self.Soil_Color_Chroma_2_CB_Hor_4, 
                        self.SoilColorValue2Completer_Hor4, self.SoilColorChroma2Completer_Hor4)

    def Hue_3_Hor_4_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_1.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_1.model().clear()
        self.Soil_Color_Value_3_CB_Hor_1.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_1.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_4.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_4, self.Soil_Color_Chroma_3_CB_Hor_4, 
                        self.SoilColorValue3Completer_Hor4, self.SoilColorChroma3Completer_Hor4)

    ### HORIZON 5 SOIL COLOR
    def Hue_1_Hor_5_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_5.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_5.model().clear()
        self.Soil_Color_Value_1_CB_Hor_5.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_5.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_5.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_5, self.Soil_Color_Chroma_1_CB_Hor_5, 
                        self.SoilColorValue1Completer_Hor5, self.SoilColorChroma1Completer_Hor5)
    
    def Hue_2_Hor_5_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_5.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_5.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_5.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_5.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_5.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_5, self.Soil_Color_Chroma_2_CB_Hor_5, 
                        self.SoilColorValue2Completer_Hor5, self.SoilColorChroma2Completer_Hor5)

    def Hue_3_Hor_5_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_5.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_5.model().clear()
        self.Soil_Color_Value_3_CB_Hor_5.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_5.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_5.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_5, self.Soil_Color_Chroma_3_CB_Hor_5, 
                        self.SoilColorValue3Completer_Hor5, self.SoilColorChroma3Completer_Hor5)

    ### HORIZON 6 SOIL COLOR
    def Hue_1_Hor_6_TextChanged(self):
        
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_1_CB_Hor_6.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_6.model().clear()
        self.Soil_Color_Value_1_CB_Hor_6.model().clear()
        self.Soil_Color_Chroma_1_CB_Hor_6.model().clear()
        currentSoilColor1Hor1 = self.Soil_Color_Hue_1_CB_Hor_6.currentText()
        munsellChange(currentSoilColor1Hor1, self.Soil_Color_Value_1_CB_Hor_6, self.Soil_Color_Chroma_1_CB_Hor_6, 
                        self.SoilColorValue1Completer_Hor6, self.SoilColorChroma1Completer_Hor6)
    
    def Hue_2_Hor_6_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_2_CB_Hor_6.currentIndex() != -1:
        # self.Mottle_Hue_1_CB_Hor_6.model().clear()
        
        self.Soil_Color_Value_2_CB_Hor_6.model().clear()
        self.Soil_Color_Chroma_2_CB_Hor_6.model().clear()
        currentSoilColor2Hor1 = self.Soil_Color_Hue_2_CB_Hor_6.currentText()
        munsellChange(currentSoilColor2Hor1, self.Soil_Color_Value_2_CB_Hor_6, self.Soil_Color_Chroma_2_CB_Hor_6, 
                        self.SoilColorValue2Completer_Hor6, self.SoilColorChroma2Completer_Hor6)

    def Hue_3_Hor_6_TextChanged(self):
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        # if self.Soil_Color_Hue_3_CB_Hor_6.currentIndex() != -1:
            
        # self.Mottle_Hue_1_CB_Hor_6.model().clear()
        self.Soil_Color_Value_3_CB_Hor_6.model().clear()
        self.Soil_Color_Chroma_3_CB_Hor_6.model().clear()
        currentSoilColor3Hor1 = self.Soil_Color_Hue_3_CB_Hor_6.currentText()
        munsellChange(currentSoilColor3Hor1, self.Soil_Color_Value_3_CB_Hor_6, self.Soil_Color_Chroma_3_CB_Hor_6, 
                        self.SoilColorValue3Completer_Hor6, self.SoilColorChroma3Completer_Hor6)


    # MOTTLE COLOR Horizon 1
    def mottleHueTextChanged_Hor1(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_1.model().clear()
        self.Mottle_Chroma_1_CB_Hor_1.model().clear()
        currentHue_Hor1 = self.Mottle_Hue_1_CB_Hor_1.currentText()
        munsellChange(currentHue_Hor1, self.Mottle_Value_1_CB_Hor_1, self.Mottle_Chroma_1_CB_Hor_1, 
                        self.mottleValueCompleter_Hor1, self.mottleChromaCompleter_Hor1)

    # MOTTLE COLOR Horizon 2
    def mottleHueTextChanged_Hor2(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_2.model().clear()
        self.Mottle_Chroma_1_CB_Hor_2.model().clear()
        currentHue_Hor2 = self.Mottle_Hue_1_CB_Hor_2.currentText()
        munsellChange(currentHue_Hor2, self.Mottle_Value_1_CB_Hor_2, self.Mottle_Chroma_1_CB_Hor_2, 
                        self.mottleValueCompleter_Hor2, self.mottleChromaCompleter_Hor2)

    # MOTTLE COLOR Horizon 3
    def mottleHueTextChanged_Hor3(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_3.model().clear()
        self.Mottle_Chroma_1_CB_Hor_3.model().clear()
        currentHue_Hor3 = self.Mottle_Hue_1_CB_Hor_3.currentText()
        munsellChange(currentHue_Hor3, self.Mottle_Value_1_CB_Hor_3, self.Mottle_Chroma_1_CB_Hor_3, 
                        self.mottleValueCompleter_Hor3, self.mottleChromaCompleter_Hor3)
    
    # MOTTLE COLOR Horizon 4
    def mottleHueTextChanged_Hor4(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_4.model().clear()
        self.Mottle_Chroma_1_CB_Hor_4.model().clear()
        currentHue_Hor4 = self.Mottle_Hue_1_CB_Hor_4.currentText()
        munsellChange(currentHue_Hor4, self.Mottle_Value_1_CB_Hor_4, self.Mottle_Chroma_1_CB_Hor_4, 
                        self.mottleValueCompleter_Hor4, self.mottleChromaCompleter_Hor4)
    
    # MOTTLE COLOR Horizon 5
    def mottleHueTextChanged_Hor5(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_5.model().clear()
        self.Mottle_Chroma_1_CB_Hor_5.model().clear()
        currentHue_Hor5 = self.Mottle_Hue_1_CB_Hor_5.currentText()
        munsellChange(currentHue_Hor5, self.Mottle_Value_1_CB_Hor_5, self.Mottle_Chroma_1_CB_Hor_5, 
                        self.mottleValueCompleter_Hor5, self.mottleChromaCompleter_Hor5)

    # MOTTLE COLOR Horizon 6                    
    def mottleHueTextChanged_Hor6(self):
        print("text changed")
        ## check if the hue combobox is not empty "but ist not working if manually typing"
        self.Mottle_Value_1_CB_Hor_6.model().clear()
        self.Mottle_Chroma_1_CB_Hor_6.model().clear()
        currentHue_Hor6 = self.Mottle_Hue_1_CB_Hor_6.currentText()
        munsellChange(currentHue_Hor6, self.Mottle_Value_1_CB_Hor_6, self.Mottle_Chroma_1_CB_Hor_6, 
                        self.mottleValueCompleter_Hor6, self.mottleChromaCompleter_Hor6)
  
    ### Focus Out for depth of soil horizon
    def horUpperFromFocusOut_Hor1(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_1.setText(self.Hor_Upper_From_LE_Hor_1.text())

    def horLowerFromFocusOut_Hor1(self):
        self.Hor_Lower_To_LE_Hor_1.setText(self.Hor_Lower_From_LE_Hor_1.text())

    def horUpperFromFocusOut_Hor2(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_2.setText(self.Hor_Upper_From_LE_Hor_2.text())

    def horLowerFromFocusOut_Hor2(self):
        self.Hor_Lower_To_LE_Hor_2.setText(self.Hor_Lower_From_LE_Hor_2.text())

    def horUpperFromFocusOut_Hor3(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_3.setText(self.Hor_Upper_From_LE_Hor_3.text())

    def horLowerFromFocusOut_Hor3(self):
        self.Hor_Lower_To_LE_Hor_3.setText(self.Hor_Lower_From_LE_Hor_3.text())

    def horUpperFromFocusOut_Hor4(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_4.setText(self.Hor_Upper_From_LE_Hor_4.text())

    def horLowerFromFocusOut_Hor4(self):
        self.Hor_Lower_To_LE_Hor_4.setText(self.Hor_Lower_From_LE_Hor_4.text())

    def horUpperFromFocusOut_Hor5(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_5.setText(self.Hor_Upper_From_LE_Hor_5.text())

    def horLowerFromFocusOut_Hor5(self):
        self.Hor_Lower_To_LE_Hor_5.setText(self.Hor_Lower_From_LE_Hor_5.text())

    def horUpperFromFocusOut_Hor6(self):
        print("focus lost") 
        self.Hor_Upper_To_LE_Hor_6.setText(self.Hor_Upper_From_LE_Hor_6.text())

    def horLowerFromFocusOut_Hor6(self):
        self.Hor_Lower_To_LE_Hor_6.setText(self.Hor_Lower_From_LE_Hor_6.text())

    ### Populate Data for Horizon
    ### Note: siteinformation is placed in populateDataHor1

    def populateDataHor1(self):

        logging.info("Horizon Windows - Populate Horizon 1")
        
        # Number Form
        # self.siteModel = QSqlTableModel()
        populate(self.siteModel, "Site", self.Number_Form_CB_Hor)

        # Copy Button and ComboBox
        self.proxyHorizonModel = QSortFilterProxyModel()
        self.proxyHorizonModel.setSourceModel(self.horizonModel)
    
        self.Form_Copy_CB.setModel(self.proxyHorizonModel)
        self.Form_Copy_CB.setModelColumn(1)
        self.Form_Copy_CB.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))
        self.Form_Copy_CB.setEnabled(False)
        self.Copy_All_PB.setEnabled(False)

        # Sorting Number Form Combobox with Ascending Order
        # self.Number_Form_CB_Hor.model().sort(1, Qt.SortOrder(Qt.AscendingOrder))
        # dont add data to noFormIndex
        self.Number_Form_CB_Hor.setInsertPolicy(self.Number_Form_CB_Hor.NoInsert)
        self.SPT_CB_Hor.setDisabled(True) 
        self.Provinsi_CB_Hor.setDisabled(True) 
        self.Kabupaten_CB_Hor.setDisabled(True) 
        self.Kecamatan_CB_Hor.setDisabled(True) 
        self.Desa_CB_Hor.setDisabled(True) 
        self.Date_CB_Hor.setDisabled(True) 
        self.Initial_Name_LE_Hor.setDisabled(True) 
        self.Observation_Number_LE_Hor.setDisabled(True) 
        self.Kind_Observation_CB_Hor.setDisabled(True) 
        self.UTM_Zone_1_LE_Hor.setDisabled(True) 
        self.UTM_Zone_2_LE_Hor.setDisabled(True) 
        self.X_East_LE_Hor.setDisabled(True) 
        self.Y_North_LE_Hor.setDisabled(True) 
        ##
        self.Number_Form_LE_Hor_1.setDisabled(True)
        self.Number_Form_LE_Hor_2.setDisabled(True)
        self.Number_Form_LE_Hor_3.setDisabled(True)
        self.Number_Form_LE_Hor_4.setDisabled(True)
        self.Number_Form_LE_Hor_5.setDisabled(True)
        self.Number_Form_LE_Hor_6.setDisabled(True)
        self.Initial_LE_Hor_1.setDisabled(True)
        self.Initial_LE_Hor_2.setDisabled(True)
        self.Initial_LE_Hor_3.setDisabled(True)
        self.Initial_LE_Hor_4.setDisabled(True)
        self.Initial_LE_Hor_5.setDisabled(True)
        self.Initial_LE_Hor_6.setDisabled(True)
        self.SPT_LE_Hor_1.setDisabled(True)
        self.SPT_LE_Hor_2.setDisabled(True)
        self.SPT_LE_Hor_3.setDisabled(True)
        self.SPT_LE_Hor_4.setDisabled(True)
        self.SPT_LE_Hor_5.setDisabled(True)
        self.SPT_LE_Hor_6.setDisabled(True)

        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_1, self.boundaryDistCompleter_Hor1, boundaryDistTableView, 80)
        
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_1, self.boundaryTopoCompleter_Hor1, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor1_Model = QSqlQueryModel()
        soilColorHue1_Hor1_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_1.setModel(soilColorHue1_Hor1_Model)
        self.Soil_Color_Hue_1_CB_Hor_1.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor1.setModel(soilColorHue1_Hor1_Model)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor1.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        self.Soil_Color_Hue_1_CB_Hor_1.setCompleter(self.SoilColorHue1Completer_Hor1)

        # SOIL COLOR 2
        soilColorHue2_Hor1_Model = QSqlQueryModel()
        soilColorHue2_Hor1_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_1.setModel(soilColorHue2_Hor1_Model)
        self.Soil_Color_Hue_2_CB_Hor_1.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor1.setModel(soilColorHue2_Hor1_Model)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor1.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        self.Soil_Color_Hue_2_CB_Hor_1.setCompleter(self.SoilColorHue2Completer_Hor1)

        # SOIL COLOR 3
        soilColorHue3_Hor1_Model = QSqlQueryModel()
        soilColorHue3_Hor1_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_1.setModel(soilColorHue3_Hor1_Model)
        self.Soil_Color_Hue_3_CB_Hor_1.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor1.setModel(soilColorHue3_Hor1_Model)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor1.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        self.Soil_Color_Hue_3_CB_Hor_1.setCompleter(self.SoilColorHue3Completer_Hor1)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_1, self.textureClassCompleter_Hor1, textureClassTableView, 110)
        
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_1, self.textureSandCompleter_Hor1, textureSandTableView, 110)
        
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_1, self.textureModSizeCompleter_Hor1, textureModSizeTableView, 110)
        
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_1, self.textureModAbundanceCompleter_Hor1, textureModAbundanceTableView, 110)
        
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_1, self.structureShape1Completer_Hor1, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_1, self.structureShape2Completer_Hor1, structureShapeTableView, 110) # structure Shape 2
        
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_1, self.structureSize1Completer_Hor1, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_1, self.structureSize2Completer_Hor1, structureSizeTableView, 80) # structure Size 2
        
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_1, self.structureGrade1Completer_Hor1, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_1, self.structureGrade2Completer_Hor1, structureGradeTableView, 80) # structure grade 2
        
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_1, self.structureRelationCompleter_Hor1, structureRelationTableView, 70)
        
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_1, self.consistenceMoistCompleter_Hor1, consistenceMoistTableView, 80)
        
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_1, self.consistenceStickinessCompleter_Hor1, consistenceStickinessTableView, 80)
        
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_1, self.consistencePlasticityCompleter_Hor1, consistencePlasticityTableView, 90)
        
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_1, self.mottleAbundanceCompleter_Hor1, mottleAbundanceTableView, 60)
        
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_1, self.mottleSizeCompleter_Hor1, mottleSizeTableView, 60)
        
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_1, self.mottleContrastCompleter_Hor1, mottleContrastTableView, 60)
        
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_1, self.mottleShapeCompleter_Hor1, mottleShapeTableView, 80)
        
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_1.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_1.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor1.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_1.setCompleter(self.mottleHueCompleter_Hor1)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_1.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_1.setCurrentIndex(-1)
        # self.SoilColorHue1Completer_Hor1.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_1.setCompleter(self.SoilColorHue1Completer_Hor1)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_1

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_1, self.rootFineCompleter_Hor1, rootFineTableView, 70)
        
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_1, self.rootMediumCompleter_Hor1, rootMediumTableView, 70)
        
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_1, self.rootCoarseCompleter_Hor1, rootCoarseTableView, 70)


        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)
    
    def populateDataHor2(self):
        logging.info("Horizon Windows - Populate Horizon 2")
   
        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_2, self.boundaryDistCompleter_Hor2, boundaryDistTableView, 80)
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_2, self.boundaryTopoCompleter_Hor2, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor2_Model = QSqlQueryModel()
        soilColorHue1_Hor2_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_2.setModel(soilColorHue1_Hor2_Model)
        self.Soil_Color_Hue_1_CB_Hor_2.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor2.setModel(soilColorHue1_Hor2_Model)
        self.Soil_Color_Hue_1_CB_Hor_2.setCompleter(self.SoilColorHue1Completer_Hor2)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor2.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 2
        soilColorHue2_Hor2_Model = QSqlQueryModel()
        soilColorHue2_Hor2_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_2.setModel(soilColorHue2_Hor2_Model)
        self.Soil_Color_Hue_2_CB_Hor_2.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor2.setModel(soilColorHue2_Hor2_Model)
        self.Soil_Color_Hue_2_CB_Hor_2.setCompleter(self.SoilColorHue2Completer_Hor2)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor2.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 3
        soilColorHue3_Hor2_Model = QSqlQueryModel()
        soilColorHue3_Hor2_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_2.setModel(soilColorHue3_Hor2_Model)
        self.Soil_Color_Hue_3_CB_Hor_2.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor2.setModel(soilColorHue3_Hor2_Model)
        self.Soil_Color_Hue_3_CB_Hor_2.setCompleter(self.SoilColorHue3Completer_Hor2)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor2.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_2, self.textureClassCompleter_Hor2, textureClassTableView, 110)
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_2, self.textureSandCompleter_Hor2, textureSandTableView, 110)
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_2, self.textureModSizeCompleter_Hor2, textureModSizeTableView, 110)
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_2, self.textureModAbundanceCompleter_Hor2, textureModAbundanceTableView, 110)
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_2, self.structureShape1Completer_Hor2, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_2, self.structureShape2Completer_Hor2, structureShapeTableView, 110) # structure Shape 2
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_2, self.structureSize1Completer_Hor2, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_2, self.structureSize2Completer_Hor2, structureSizeTableView, 80) # structure Size 2
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_2, self.structureGrade1Completer_Hor2, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_2, self.structureGrade2Completer_Hor2, structureGradeTableView, 80) # structure grade 2
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_2, self.structureRelationCompleter_Hor2, structureRelationTableView, 70)
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_2, self.consistenceMoistCompleter_Hor2, consistenceMoistTableView, 80)
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_2, self.consistenceStickinessCompleter_Hor2, consistenceStickinessTableView, 80)
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_2, self.consistencePlasticityCompleter_Hor2, consistencePlasticityTableView, 90)
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_2, self.mottleAbundanceCompleter_Hor2, mottleAbundanceTableView, 60)
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_2, self.mottleSizeCompleter_Hor2, mottleSizeTableView, 60)
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_2, self.mottleContrastCompleter_Hor2, mottleContrastTableView, 60)
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_2, self.mottleShapeCompleter_Hor2, mottleShapeTableView, 80)
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_2.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_2.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor2.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_2.setCompleter(self.mottleHueCompleter_Hor2)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_2.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_2.setCurrentIndex(-1)
        # self.soilColorHue1Completer_Hor2.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_2.setCompleter(self.soilColorHue1Completer_Hor2)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_2

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_2, self.rootFineCompleter_Hor2, rootFineTableView, 70)
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_2, self.rootMediumCompleter_Hor2, rootMediumTableView, 70)
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_2, self.rootCoarseCompleter_Hor2, rootCoarseTableView, 70)

        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)

    def populateDataHor3(self):
        logging.info("Horizon Windows - Populate Horizon 3")

        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_3, self.boundaryDistCompleter_Hor3, boundaryDistTableView, 80)
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_3, self.boundaryTopoCompleter_Hor3, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor3_Model = QSqlQueryModel()
        soilColorHue1_Hor3_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_3.setModel(soilColorHue1_Hor3_Model)
        self.Soil_Color_Hue_1_CB_Hor_3.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor3.setModel(soilColorHue1_Hor3_Model)
        self.Soil_Color_Hue_1_CB_Hor_3.setCompleter(self.SoilColorHue1Completer_Hor3)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor3.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 2
        soilColorHue2_Hor3_Model = QSqlQueryModel()
        soilColorHue2_Hor3_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_3.setModel(soilColorHue2_Hor3_Model)
        self.Soil_Color_Hue_2_CB_Hor_3.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor3.setModel(soilColorHue2_Hor3_Model)
        self.Soil_Color_Hue_2_CB_Hor_3.setCompleter(self.SoilColorHue2Completer_Hor3)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor3.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        
        # SOIL COLOR 3
        soilColorHue3_Hor3_Model = QSqlQueryModel()
        soilColorHue3_Hor3_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_3.setModel(soilColorHue3_Hor3_Model)
        self.Soil_Color_Hue_3_CB_Hor_3.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor3.setModel(soilColorHue3_Hor3_Model)
        self.Soil_Color_Hue_3_CB_Hor_3.setCompleter(self.SoilColorHue3Completer_Hor3)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor3.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_3, self.textureClassCompleter_Hor3, textureClassTableView, 110)
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_3, self.textureSandCompleter_Hor3, textureSandTableView, 110)
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_3, self.textureModSizeCompleter_Hor3, textureModSizeTableView, 110)
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_3, self.textureModAbundanceCompleter_Hor3, textureModAbundanceTableView, 110)
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_3, self.structureShape1Completer_Hor3, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_3, self.structureShape2Completer_Hor3, structureShapeTableView, 110) # structure Shape 2
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_3, self.structureSize1Completer_Hor3, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_3, self.structureSize2Completer_Hor3, structureSizeTableView, 80) # structure Size 2
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_3, self.structureGrade1Completer_Hor3, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_3, self.structureGrade2Completer_Hor3, structureGradeTableView, 80) # structure grade 2
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_3, self.structureRelationCompleter_Hor3, structureRelationTableView, 70)
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_3, self.consistenceMoistCompleter_Hor3, consistenceMoistTableView, 80)
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_3, self.consistenceStickinessCompleter_Hor3, consistenceStickinessTableView, 80)
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_3, self.consistencePlasticityCompleter_Hor3, consistencePlasticityTableView, 90)
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_3, self.mottleAbundanceCompleter_Hor3, mottleAbundanceTableView, 60)
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_3, self.mottleSizeCompleter_Hor3, mottleSizeTableView, 60)
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_3, self.mottleContrastCompleter_Hor3, mottleContrastTableView, 60)
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_3, self.mottleShapeCompleter_Hor3, mottleShapeTableView, 80)
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_3.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_3.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor3.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_3.setCompleter(self.mottleHueCompleter_Hor3)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_3.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_3.setCurrentIndex(-1)
        # self.soilColorHue1Completer_Hor3.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_3.setCompleter(self.soilColorHue1Completer_Hor3)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_3

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_3, self.rootFineCompleter_Hor3, rootFineTableView, 70)
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_3, self.rootMediumCompleter_Hor3, rootMediumTableView, 70)
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_3, self.rootCoarseCompleter_Hor3, rootCoarseTableView, 70)
       
        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)

    def populateDataHor4(self):
        logging.info("Horizon Windows - Populate Horizon 4")

        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_4, self.boundaryDistCompleter_Hor4, boundaryDistTableView, 80)
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_4, self.boundaryTopoCompleter_Hor4, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor4_Model = QSqlQueryModel()
        soilColorHue1_Hor4_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_4.setModel(soilColorHue1_Hor4_Model)
        self.Soil_Color_Hue_1_CB_Hor_4.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor4.setModel(soilColorHue1_Hor4_Model)
        self.Soil_Color_Hue_1_CB_Hor_4.setCompleter(self.SoilColorHue1Completer_Hor4)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor4.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 2
        soilColorHue2_Hor4_Model = QSqlQueryModel()
        soilColorHue2_Hor4_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_4.setModel(soilColorHue2_Hor4_Model)
        self.Soil_Color_Hue_2_CB_Hor_4.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor4.setModel(soilColorHue2_Hor4_Model)
        self.Soil_Color_Hue_2_CB_Hor_4.setCompleter(self.SoilColorHue2Completer_Hor4)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor4.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 3
        soilColorHue3_Hor4_Model = QSqlQueryModel()
        soilColorHue3_Hor4_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_4.setModel(soilColorHue3_Hor4_Model)
        self.Soil_Color_Hue_3_CB_Hor_4.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor4.setModel(soilColorHue3_Hor4_Model)
        self.Soil_Color_Hue_3_CB_Hor_4.setCompleter(self.SoilColorHue3Completer_Hor4)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor4.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_4, self.textureClassCompleter_Hor4, textureClassTableView, 110)
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_4, self.textureSandCompleter_Hor4, textureSandTableView, 110)
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_4, self.textureModSizeCompleter_Hor4, textureModSizeTableView, 110)
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_4, self.textureModAbundanceCompleter_Hor4, textureModAbundanceTableView, 110)
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_4, self.structureShape1Completer_Hor4, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_4, self.structureShape2Completer_Hor4, structureShapeTableView, 110) # structure Shape 2
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_4, self.structureSize1Completer_Hor4, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_4, self.structureSize2Completer_Hor4, structureSizeTableView, 80) # structure Size 2
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_4, self.structureGrade1Completer_Hor4, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_4, self.structureGrade2Completer_Hor4, structureGradeTableView, 80) # structure grade 2
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_4, self.structureRelationCompleter_Hor4, structureRelationTableView, 70)
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_4, self.consistenceMoistCompleter_Hor4, consistenceMoistTableView, 80)
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_4, self.consistenceStickinessCompleter_Hor4, consistenceStickinessTableView, 80)
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_4, self.consistencePlasticityCompleter_Hor4, consistencePlasticityTableView, 90)
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_4, self.mottleAbundanceCompleter_Hor4, mottleAbundanceTableView, 60)
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_4, self.mottleSizeCompleter_Hor4, mottleSizeTableView, 60)
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_4, self.mottleContrastCompleter_Hor4, mottleContrastTableView, 60)
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_4, self.mottleShapeCompleter_Hor4, mottleShapeTableView, 80)
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_4.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_4.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor4.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_4.setCompleter(self.mottleHueCompleter_Hor4)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_4.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_4.setCurrentIndex(-1)
        # self.soilColorHue1Completer_Hor4.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_4.setCompleter(self.soilColorHue1Completer_Hor4)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_4

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_4, self.rootFineCompleter_Hor4, rootFineTableView, 70)
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_4, self.rootMediumCompleter_Hor4, rootMediumTableView, 70)
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_4, self.rootCoarseCompleter_Hor4, rootCoarseTableView, 70)
        
        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)

    def populateDataHor5(self):
        logging.info("Horizon Windows - Populate Horizon 5")

        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_5, self.boundaryDistCompleter_Hor5, boundaryDistTableView, 80)
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_5, self.boundaryTopoCompleter_Hor5, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor5_Model = QSqlQueryModel()
        soilColorHue1_Hor5_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_5.setModel(soilColorHue1_Hor5_Model)
        self.Soil_Color_Hue_1_CB_Hor_5.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor5.setModel(soilColorHue1_Hor5_Model)
        self.Soil_Color_Hue_1_CB_Hor_5.setCompleter(self.SoilColorHue1Completer_Hor5)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor5.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 2
        soilColorHue2_Hor5_Model = QSqlQueryModel()
        soilColorHue2_Hor5_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_5.setModel(soilColorHue2_Hor5_Model)
        self.Soil_Color_Hue_2_CB_Hor_5.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor5.setModel(soilColorHue2_Hor5_Model)
        self.Soil_Color_Hue_2_CB_Hor_5.setCompleter(self.SoilColorHue2Completer_Hor5)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor5.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 3
        soilColorHue3_Hor5_Model = QSqlQueryModel()
        soilColorHue3_Hor5_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_5.setModel(soilColorHue3_Hor5_Model)
        self.Soil_Color_Hue_3_CB_Hor_5.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor5.setModel(soilColorHue3_Hor5_Model)
        self.Soil_Color_Hue_3_CB_Hor_5.setCompleter(self.SoilColorHue3Completer_Hor5)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor5.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_5, self.textureClassCompleter_Hor5, textureClassTableView, 110)
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_5, self.textureSandCompleter_Hor5, textureSandTableView, 110)
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_5, self.textureModSizeCompleter_Hor5, textureModSizeTableView, 110)
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_5, self.textureModAbundanceCompleter_Hor5, textureModAbundanceTableView, 110)
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_5, self.structureShape1Completer_Hor5, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_5, self.structureShape2Completer_Hor5, structureShapeTableView, 110) # structure Shape 2
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_5, self.structureSize1Completer_Hor5, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_5, self.structureSize2Completer_Hor5, structureSizeTableView, 80) # structure Size 2
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_5, self.structureGrade1Completer_Hor5, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_5, self.structureGrade2Completer_Hor5, structureGradeTableView, 80) # structure grade 2
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_5, self.structureRelationCompleter_Hor5, structureRelationTableView, 70)
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_5, self.consistenceMoistCompleter_Hor5, consistenceMoistTableView, 80)
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_5, self.consistenceStickinessCompleter_Hor5, consistenceStickinessTableView, 80)
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_5, self.consistencePlasticityCompleter_Hor5, consistencePlasticityTableView, 90)
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_5, self.mottleAbundanceCompleter_Hor5, mottleAbundanceTableView, 60)
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_5, self.mottleSizeCompleter_Hor5, mottleSizeTableView, 60)
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_5, self.mottleContrastCompleter_Hor5, mottleContrastTableView, 60)
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_5, self.mottleShapeCompleter_Hor5, mottleShapeTableView, 80)
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_5.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_5.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor5.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_5.setCompleter(self.mottleHueCompleter_Hor5)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_5.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_5.setCurrentIndex(-1)
        # self.soilColorHue1Completer_Hor5.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_5.setCompleter(self.soilColorHue1Completer_Hor5)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_5

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_5, self.rootFineCompleter_Hor5, rootFineTableView, 70)
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_5, self.rootMediumCompleter_Hor5, rootMediumTableView, 70)
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_5, self.rootCoarseCompleter_Hor5, rootCoarseTableView, 70)
        
        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)

    def populateDataHor6(self):  
        logging.info("Horizon Windows - Populate Horizon 6")

        # Boundary Distribution
        boundaryDistModel = QSqlTableModel()
        boundaryDistTableView = QTableView()
        populatePlusTable(boundaryDistModel, "BoundaryDist", self.Boundary_Dist_CB_Hor_6, self.boundaryDistCompleter_Hor6, boundaryDistTableView, 80)
        # Boundary Topo
        boundaryTopoModel = QSqlTableModel()
        boundaryTopoTableView = QTableView()
        populatePlusTable(boundaryTopoModel, "BoundaryTopo", self.Boundary_Topo_CB_Hor_6, self.boundaryTopoCompleter_Hor6, boundaryTopoTableView, 80)
        
        # Soil Color Hue is using QSqlQueryModel different from the others 
        # because QSqlTableModel does not have "distint" function in query

        # SOIL COLOR 1
        soilColorHue1_Hor6_Model = QSqlQueryModel()
        soilColorHue1_Hor6_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_1_CB_Hor_6.setModel(soilColorHue1_Hor6_Model)
        self.Soil_Color_Hue_1_CB_Hor_6.setCurrentIndex(-1)
        self.SoilColorHue1Completer_Hor6.setModel(soilColorHue1_Hor6_Model)
        self.Soil_Color_Hue_1_CB_Hor_6.setCompleter(self.SoilColorHue1Completer_Hor6)
        # set the completer case insensitive
        self.SoilColorHue1Completer_Hor6.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 2
        soilColorHue2_Hor6_Model = QSqlQueryModel()
        soilColorHue2_Hor6_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_2_CB_Hor_6.setModel(soilColorHue2_Hor6_Model)
        self.Soil_Color_Hue_2_CB_Hor_6.setCurrentIndex(-1)
        self.SoilColorHue2Completer_Hor6.setModel(soilColorHue2_Hor6_Model)
        self.Soil_Color_Hue_2_CB_Hor_6.setCompleter(self.SoilColorHue2Completer_Hor6)
        # set the completer case insensitive
        self.SoilColorHue2Completer_Hor6.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # SOIL COLOR 3
        soilColorHue3_Hor6_Model = QSqlQueryModel()
        soilColorHue3_Hor6_Model.setQuery('select distinct HUE from SoilColor')
        self.Soil_Color_Hue_3_CB_Hor_6.setModel(soilColorHue3_Hor6_Model)
        self.Soil_Color_Hue_3_CB_Hor_6.setCurrentIndex(-1)
        self.SoilColorHue3Completer_Hor6.setModel(soilColorHue3_Hor6_Model)
        self.Soil_Color_Hue_3_CB_Hor_6.setCompleter(self.SoilColorHue3Completer_Hor6)
        # set the completer case insensitive
        self.SoilColorHue3Completer_Hor6.setCaseSensitivity(QtCore.Qt.CaseInsensitive)

        # Texture Class
        textureClassModel = QSqlTableModel()
        textureClassTableView = QTableView()
        populatePlusTable(textureClassModel, "TextureClass", self.Texture_Class_CB_Hor_6, self.textureClassCompleter_Hor6, textureClassTableView, 110)
        # Texture Sand
        textureSandModel = QSqlTableModel()
        textureSandTableView = QTableView()
        populatePlusTable(textureSandModel, "TextureSand", self.Texture_Sand_CB_Hor_6, self.textureSandCompleter_Hor6, textureSandTableView, 110)
        #Texture Modifier Size
        textureModSizeModel = QSqlTableModel()
        textureModSizeTableView = QTableView()
        populatePlusTable(textureModSizeModel, "TextureModSize", self.Texture_Mod_Size_CB_Hor_6, self.textureModSizeCompleter_Hor6, textureModSizeTableView, 110)
        #Texture Modifier Abundance
        textureModAbundanceModel = QSqlTableModel()
        textureModAbundanceTableView = QTableView()
        populatePlusTable(textureModAbundanceModel, "TextureModAbundance", self.Texture_Mod_Abundance_CB_Hor_6, self.textureModAbundanceCompleter_Hor6, textureModAbundanceTableView, 110)
        # Structure Shape 
        structureShapeModel = QSqlTableModel()
        structureShape2Model = QSqlTableModel()
        structureShapeTableView = QTableView()
        populatePlusTable(structureShapeModel, "StructureShape", self.Structure_Shape_1_CB_Hor_6, self.structureShape1Completer_Hor6, structureShapeTableView, 110) # structure Shape 1
        populatePlusTable(structureShape2Model, "StructureShape", self.Structure_Shape_2_CB_Hor_6, self.structureShape2Completer_Hor6, structureShapeTableView, 110) # structure Shape 2
        # Structure Size
        structureSizeModel = QSqlTableModel()
        structureSize2Model = QSqlTableModel()
        structureSizeTableView = QTableView()
        populatePlusTable(structureSizeModel, "StructureSize", self.Structure_Size_1_CB_Hor_6, self.structureSize1Completer_Hor6, structureSizeTableView, 80) # structure Size 1
        populatePlusTable(structureSize2Model, "StructureSize", self.Structure_Size_2_CB_Hor_6, self.structureSize2Completer_Hor6, structureSizeTableView, 80) # structure Size 2
        # Structure Grade
        structureGradeModel = QSqlTableModel()
        structureGrade2Model = QSqlTableModel()
        structureGradeTableView = QTableView()
        populatePlusTable(structureGradeModel, "StructureGrade", self.Structure_Grade_1_CB_Hor_6, self.structureGrade1Completer_Hor6, structureGradeTableView, 80) # structure grade 1
        populatePlusTable(structureGrade2Model, "StructureGrade", self.Structure_Grade_2_CB_Hor_6, self.structureGrade2Completer_Hor6, structureGradeTableView, 80) # structure grade 2
        # Structure Relation
        structureRelationModel = QSqlTableModel()
        structureRelationTableView = QTableView()
        populatePlusTable(structureRelationModel, "StructureRelation", self.Structure_Relation_CB_Hor_6, self.structureRelationCompleter_Hor6, structureRelationTableView, 70)
        # Consistence Moist
        consistenceMoistModel = QSqlTableModel()
        consistenceMoistTableView = QTableView()
        populatePlusTable(consistenceMoistModel, "ConsistenceMoist", self.Consistence_Moist_CB_Hor_6, self.consistenceMoistCompleter_Hor6, consistenceMoistTableView, 80)
        # Consistence Stickiness
        consistenceStickinessModel = QSqlTableModel()
        consistenceStickinessTableView = QTableView()
        populatePlusTable(consistenceStickinessModel, "ConsistenceStickiness", self.Consistence_Stickiness_CB_Hor_6, self.consistenceStickinessCompleter_Hor6, consistenceStickinessTableView, 80)
        # Consistence Plasticity
        consistencePlasticityModel = QSqlTableModel()
        consistencePlasticityTableView = QTableView()
        populatePlusTable(consistencePlasticityModel, "ConsistencePlasticity", self.Consistence_Plasticity_CB_Hor_6, self.consistencePlasticityCompleter_Hor6, consistencePlasticityTableView, 90)
        # Mottle Abundance
        mottleAbundanceModel = QSqlTableModel()
        mottleAbundanceTableView = QTableView()
        populatePlusTable(mottleAbundanceModel, "MottlesAbundance", self.Mottle_Abundance_1_CB_Hor_6, self.mottleAbundanceCompleter_Hor6, mottleAbundanceTableView, 60)
        # Mottle Size
        mottleSizeModel = QSqlTableModel()
        mottleSizeTableView = QTableView()
        populatePlusTable(mottleSizeModel, "MottlesSize", self.Mottle_Size_1_CB_Hor_6, self.mottleSizeCompleter_Hor6, mottleSizeTableView, 60)
        # Mottle Contrast
        mottleContrastModel = QSqlTableModel()
        mottleContrastTableView = QTableView()
        populatePlusTable(mottleContrastModel, "MottlesContrast", self.Mottle_Contrast_1_CB_Hor_6, self.mottleContrastCompleter_Hor6, mottleContrastTableView, 60)
        # Mottle Shape
        mottleShapeModel = QSqlTableModel()
        mottleShapeTableView = QTableView()
        populatePlusTable(mottleShapeModel, "MottlesShape", self.Mottle_Shape_1_CB_Hor_6, self.mottleShapeCompleter_Hor6, mottleShapeTableView, 80)
        # Mottle Hue is using QSqlQueryModel different from the others
        mottleHueModel = QSqlQueryModel()
        mottleHueModel.setQuery('select distinct HUE from SoilColor')
        self.Mottle_Hue_1_CB_Hor_6.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_6.setCurrentIndex(-1)
        self.mottleHueCompleter_Hor6.setModel(mottleHueModel)
        self.Mottle_Hue_1_CB_Hor_6.setCompleter(self.mottleHueCompleter_Hor6)

        # # SOIL COLOR QLINEEDIT
        # # self.Soil_Color_Hue_1_CB_Hor_6. setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_6.setCurrentIndex(-1)
        # self.soilColorHue1Completer_Hor6.setModel(mottleHueModel)
        # self.Soil_Color_Hue_1_CB_Hor_6.setCompleter(self.soilColorHue1Completer_Hor6)
            
        ### Mottle Value and Chroma not needed to be populated because
        ### already done using textChange function in Mottle_Hue_1_CB_Hor_6

        # Root Fine
        rootFineModel = QSqlTableModel()
        rootFineTableView = QTableView()
        populatePlusTable(rootFineModel, "RootsFine", self.Root_Fine_CB_Hor_6, self.rootFineCompleter_Hor6, rootFineTableView, 70)
        # Root Medium
        rootMediumModel = QSqlTableModel()
        rootMediumTableView = QTableView()
        populatePlusTable(rootMediumModel, "RootsMedium", self.Root_Medium_CB_Hor_6, self.rootMediumCompleter_Hor6, rootMediumTableView, 70)
        # Root Coarse
        rootCoarseModel = QSqlTableModel()
        rootCoarseTableView = QTableView()
        populatePlusTable(rootCoarseModel, "RootsCoarse", self.Root_Coarse_CB_Hor_6, self.rootCoarseCompleter_Hor6, rootCoarseTableView, 70)
        
        self.Next_PB.setEnabled(False)
        self.Previous_PB.setEnabled(False)

    def setTableViewCombobox(self):
        logging.info("Horizon Windows - Set Table view combobox")

        # self.textureClassTableView = QTableView()
        # tableViewCombobox(self.textureClassTableView, self.Texture_Class_CB_Hor_1, 120)
        # self.textureSandTableView = QTableView()
        # tableViewCombobox(self.textureSandTableView, self.Texture_Sand_CB_Hor_1, 120)
        # self.textureModSizeTableView = QTableView()
        # tableViewCombobox(self.textureModSizeTableView, self.Texture_Mod_Size_CB_Hor_1, 120)
        # self.textureModAbundanceTableView = QTableView()
        # tableViewCombobox(self.textureModAbundanceTableView, self.Texture_Mod_Abundance_CB_Hor_1, 120)
        # self.structureShapeTableView = QTableView()
        # tableViewCombobox(self.structureShapeTableView, T)
       
    def initMapper(self):
        logging.info("Horizon Windows - Initialize Mapper")

        # Set The QDataWidgetMapper
        self.horizonMapper = QDataWidgetMapper()
        self.horizonMapper.setModel(self.horizonModel)
        self.horizonMapper.setSubmitPolicy(QDataWidgetMapper.ManualSubmit)

        # TODO: still not working properlly need to check
        self.siteMapper = QDataWidgetMapper()
        self.siteMapper.setModel(self.siteModel)
        self.siteMapper.setSubmitPolicy(QDataWidgetMapper.ManualSubmit)

        ### Site Information Mapping
        self.siteMapper.addMapping(self.Number_Form_CB_Hor, 1)
        self.siteMapper.addMapping(self.SPT_CB_Hor, 2)
        self.siteMapper.addMapping(self.Provinsi_CB_Hor, 3)
        self.siteMapper.addMapping(self.Kabupaten_CB_Hor, 4)
        self.siteMapper.addMapping(self.Kecamatan_CB_Hor, 5)
        self.siteMapper.addMapping(self.Desa_CB_Hor, 6)
        self.siteMapper.addMapping(self.Date_CB_Hor, 7)
        self.siteMapper.addMapping(self.Initial_Name_LE_Hor, 8)
        self.siteMapper.addMapping(self.Observation_Number_LE_Hor, 9)
        self.siteMapper.addMapping(self.Kind_Observation_CB_Hor, 10)
        self.siteMapper.addMapping(self.UTM_Zone_1_LE_Hor, 11)
        self.siteMapper.addMapping(self.UTM_Zone_2_LE_Hor, 12)
        self.siteMapper.addMapping(self.X_East_LE_Hor, 13)
        self.siteMapper.addMapping(self.Y_North_LE_Hor, 14)

        ### Horizon 1 mapping
        ### It seems working properly for the others horizon, I dont know what
        ### the catch is. Should I still add mapper to 2-6 horizons???
        self.horizonMapper.addMapping(self.Number_Form_CB_Hor, 1)
        self.horizonMapper.addMapping(self.Hor_Design_Discon_LE_Hor_1, 2)
        self.horizonMapper.addMapping(self.Hor_Design_Master_LE_Hor_1, 3)
        self.horizonMapper.addMapping(self.Hor_Design_Sub_LE_Hor_1, 4)
        self.horizonMapper.addMapping(self.Hor_Design_Number_LE_Hor_1, 5)
        self.horizonMapper.addMapping(self.Hor_Upper_From_LE_Hor_1, 6)
        self.horizonMapper.addMapping(self.Hor_Upper_To_LE_Hor_1, 7)
        self.horizonMapper.addMapping(self.Hor_Lower_From_LE_Hor_1, 8)
        self.horizonMapper.addMapping(self.Hor_Lower_To_LE_Hor_1, 9)
        self.horizonMapper.addMapping(self.Boundary_Dist_CB_Hor_1, 10)
        self.horizonMapper.addMapping(self.Boundary_Topo_CB_Hor_1, 11)
        # self.horizonMapper.addMapping(self.Soil_Color_1_Moisture, 12)
        self.horizonMapper.addMapping(self.Soil_Color_Hue_1_CB_Hor_1, 13)
        self.horizonMapper.addMapping(self.Soil_Color_Value_1_CB_Hor_1, 14)
        self.horizonMapper.addMapping(self.Soil_Color_Chroma_1_CB_Hor_1, 15)
        # self.horizonMapper.addMapping(self.Soil_Color_2_Moisture, 16)
        self.horizonMapper.addMapping(self.Soil_Color_Hue_2_CB_Hor_1, 17)
        self.horizonMapper.addMapping(self.Soil_Color_Value_2_CB_Hor_1, 18)
        self.horizonMapper.addMapping(self.Soil_Color_Chroma_2_CB_Hor_1, 19)
        # self.horizonMapper.addMapping(self.Soil_Color_3_Moisture, 20)
        self.horizonMapper.addMapping(self.Soil_Color_Hue_3_CB_Hor_1, 21)
        self.horizonMapper.addMapping(self.Soil_Color_Value_3_CB_Hor_1, 22)
        self.horizonMapper.addMapping(self.Soil_Color_Chroma_3_CB_Hor_1, 23)
        self.horizonMapper.addMapping(self.Texture_Class_CB_Hor_1, 24)
        self.horizonMapper.addMapping(self.Texture_Sand_CB_Hor_1, 25)
        self.horizonMapper.addMapping(self.Texture_Mod_Size_CB_Hor_1, 26)
        self.horizonMapper.addMapping(self.Texture_Mod_Abundance_CB_Hor_1, 27)
        self.horizonMapper.addMapping(self.Structure_Shape_1_CB_Hor_1, 28)
        self.horizonMapper.addMapping(self.Structure_Size_1_CB_Hor_1, 29)
        self.horizonMapper.addMapping(self.Structure_Grade_1_CB_Hor_1, 30)
        self.horizonMapper.addMapping(self.Structure_Relation_CB_Hor_1, 31)
        self.horizonMapper.addMapping(self.Structure_Shape_2_CB_Hor_1, 32)
        self.horizonMapper.addMapping(self.Structure_Size_2_CB_Hor_1, 33)
        self.horizonMapper.addMapping(self.Structure_Grade_2_CB_Hor_1, 34)
        self.horizonMapper.addMapping(self.Consistence_Moist_CB_Hor_1, 35)
        self.horizonMapper.addMapping(self.Consistence_Stickiness_CB_Hor_1, 36)
        self.horizonMapper.addMapping(self.Consistence_Plasticity_CB_Hor_1, 37)
        self.horizonMapper.addMapping(self.Mottle_Abundance_1_CB_Hor_1, 38)
        self.horizonMapper.addMapping(self.Mottle_Size_1_CB_Hor_1, 39)
        self.horizonMapper.addMapping(self.Mottle_Contrast_1_CB_Hor_1, 40)
        self.horizonMapper.addMapping(self.Mottle_Shape_1_CB_Hor_1, 41)
        self.horizonMapper.addMapping(self.Mottle_Hue_1_CB_Hor_1, 42)
        self.horizonMapper.addMapping(self.Mottle_Value_1_CB_Hor_1, 43)
        self.horizonMapper.addMapping(self.Mottle_Chroma_1_CB_Hor_1, 44)
        self.horizonMapper.addMapping(self.Root_Fine_CB_Hor_1, 45)
        self.horizonMapper.addMapping(self.Root_Medium_CB_Hor_1, 46)
        self.horizonMapper.addMapping(self.Root_Coarse_CB_Hor_1, 47)

        # self.horizonMapper.toLast()

    def initCompleter(self):
        logging.info("Horizon Windows - Initialize Completer")

        self.noFormHorizonCompleter = QCompleter()

        self.noFormCompleter = QCompleter()
        
        ### Horizon 1
        self.boundaryDistCompleter_Hor1 = QCompleter()
        self.boundaryTopoCompleter_Hor1 = QCompleter()
        self.SoilColorHue1Completer_Hor1 = QCompleter()
        self.SoilColorValue1Completer_Hor1 = QCompleter()
        self.SoilColorChroma1Completer_Hor1 = QCompleter()
        self.SoilColorHue2Completer_Hor1 = QCompleter()
        self.SoilColorValue2Completer_Hor1 = QCompleter()
        self.SoilColorChroma2Completer_Hor1 = QCompleter()
        self.SoilColorHue3Completer_Hor1 = QCompleter()
        self.SoilColorValue3Completer_Hor1 = QCompleter()
        self.SoilColorChroma3Completer_Hor1 = QCompleter()
        self.textureClassCompleter_Hor1 = QCompleter()
        self.textureSandCompleter_Hor1 = QCompleter()
        self.textureModSizeCompleter_Hor1 = QCompleter()
        self.textureModAbundanceCompleter_Hor1 = QCompleter()
        self.structureShape1Completer_Hor1 = QCompleter()
        self.structureSize1Completer_Hor1 = QCompleter()
        self.structureGrade1Completer_Hor1 = QCompleter()
        self.structureRelationCompleter_Hor1 = QCompleter()
        self.structureShape2Completer_Hor1 = QCompleter()
        self.structureSize2Completer_Hor1 = QCompleter()
        self.structureGrade2Completer_Hor1 = QCompleter()
        self.consistenceMoistCompleter_Hor1 = QCompleter()
        self.consistenceStickinessCompleter_Hor1 = QCompleter()
        self.consistencePlasticityCompleter_Hor1 = QCompleter()
        self.mottleAbundanceCompleter_Hor1 = QCompleter()
        self.mottleSizeCompleter_Hor1 = QCompleter()
        self.mottleContrastCompleter_Hor1 = QCompleter()
        self.mottleShapeCompleter_Hor1 = QCompleter()
        self.mottleHueCompleter_Hor1 = QCompleter()
        self.mottleValueCompleter_Hor1 = QCompleter()
        self.mottleChromaCompleter_Hor1 = QCompleter()
        self.rootFineCompleter_Hor1 = QCompleter()
        self.rootMediumCompleter_Hor1 = QCompleter()
        self.rootCoarseCompleter_Hor1 = QCompleter()

        ### Horizon 2
        self.boundaryDistCompleter_Hor2 = QCompleter()
        self.boundaryTopoCompleter_Hor2 = QCompleter()
        self.SoilColorHue1Completer_Hor2 = QCompleter()
        self.SoilColorValue1Completer_Hor2 = QCompleter()
        self.SoilColorChroma1Completer_Hor2 = QCompleter()
        self.SoilColorHue2Completer_Hor2 = QCompleter()
        self.SoilColorValue2Completer_Hor2 = QCompleter()
        self.SoilColorChroma2Completer_Hor2 = QCompleter()
        self.SoilColorHue3Completer_Hor2 = QCompleter()
        self.SoilColorValue3Completer_Hor2 = QCompleter()
        self.SoilColorChroma3Completer_Hor2 = QCompleter()
        self.textureClassCompleter_Hor2 = QCompleter()
        self.textureSandCompleter_Hor2 = QCompleter()
        self.textureModSizeCompleter_Hor2 = QCompleter()
        self.textureModAbundanceCompleter_Hor2 = QCompleter()
        self.structureShape1Completer_Hor2 = QCompleter()
        self.structureSize1Completer_Hor2 = QCompleter()
        self.structureGrade1Completer_Hor2 = QCompleter()
        self.structureRelationCompleter_Hor2 = QCompleter()
        self.structureShape2Completer_Hor2 = QCompleter()
        self.structureSize2Completer_Hor2 = QCompleter()
        self.structureGrade2Completer_Hor2 = QCompleter()
        self.consistenceMoistCompleter_Hor2 = QCompleter()
        self.consistenceStickinessCompleter_Hor2 = QCompleter()
        self.consistencePlasticityCompleter_Hor2 = QCompleter()
        self.mottleAbundanceCompleter_Hor2 = QCompleter()
        self.mottleSizeCompleter_Hor2 = QCompleter()
        self.mottleContrastCompleter_Hor2 = QCompleter()
        self.mottleShapeCompleter_Hor2 = QCompleter()
        self.mottleHueCompleter_Hor2 = QCompleter()
        self.mottleValueCompleter_Hor2 = QCompleter()
        self.mottleChromaCompleter_Hor2 = QCompleter()
        self.rootFineCompleter_Hor2 = QCompleter()
        self.rootMediumCompleter_Hor2 = QCompleter()
        self.rootCoarseCompleter_Hor2 = QCompleter()

        ### Horizon 3
        self.boundaryDistCompleter_Hor3 = QCompleter()
        self.boundaryTopoCompleter_Hor3 = QCompleter()
        self.SoilColorHue1Completer_Hor3 = QCompleter()
        self.SoilColorValue1Completer_Hor3 = QCompleter()
        self.SoilColorChroma1Completer_Hor3 = QCompleter()
        self.SoilColorHue2Completer_Hor3 = QCompleter()
        self.SoilColorValue2Completer_Hor3 = QCompleter()
        self.SoilColorChroma2Completer_Hor3 = QCompleter()
        self.SoilColorHue3Completer_Hor3 = QCompleter()
        self.SoilColorValue3Completer_Hor3 = QCompleter()
        self.SoilColorChroma3Completer_Hor3 = QCompleter()
        self.textureClassCompleter_Hor3 = QCompleter()
        self.textureSandCompleter_Hor3 = QCompleter()
        self.textureModSizeCompleter_Hor3 = QCompleter()
        self.textureModAbundanceCompleter_Hor3 = QCompleter()
        self.structureShape1Completer_Hor3 = QCompleter()
        self.structureSize1Completer_Hor3 = QCompleter()
        self.structureGrade1Completer_Hor3 = QCompleter()
        self.structureRelationCompleter_Hor3 = QCompleter()
        self.structureShape2Completer_Hor3 = QCompleter()
        self.structureSize2Completer_Hor3 = QCompleter()
        self.structureGrade2Completer_Hor3 = QCompleter()
        self.consistenceMoistCompleter_Hor3 = QCompleter()
        self.consistenceStickinessCompleter_Hor3 = QCompleter()
        self.consistencePlasticityCompleter_Hor3 = QCompleter()
        self.mottleAbundanceCompleter_Hor3 = QCompleter()
        self.mottleSizeCompleter_Hor3 = QCompleter()
        self.mottleContrastCompleter_Hor3 = QCompleter()
        self.mottleShapeCompleter_Hor3 = QCompleter()
        self.mottleHueCompleter_Hor3 = QCompleter()
        self.mottleValueCompleter_Hor3 = QCompleter()
        self.mottleChromaCompleter_Hor3 = QCompleter()
        self.rootFineCompleter_Hor3 = QCompleter()
        self.rootMediumCompleter_Hor3 = QCompleter()
        self.rootCoarseCompleter_Hor3 = QCompleter()

        ### Horizon 4
        self.boundaryDistCompleter_Hor4 = QCompleter()
        self.boundaryTopoCompleter_Hor4 = QCompleter()
        self.SoilColorHue1Completer_Hor4 = QCompleter()
        self.SoilColorValue1Completer_Hor4 = QCompleter()
        self.SoilColorChroma1Completer_Hor4 = QCompleter()
        self.SoilColorHue2Completer_Hor4 = QCompleter()
        self.SoilColorValue2Completer_Hor4 = QCompleter()
        self.SoilColorChroma2Completer_Hor4 = QCompleter()
        self.SoilColorHue3Completer_Hor4 = QCompleter()
        self.SoilColorValue3Completer_Hor4 = QCompleter()
        self.SoilColorChroma3Completer_Hor4 = QCompleter()
        self.textureClassCompleter_Hor4 = QCompleter()
        self.textureSandCompleter_Hor4 = QCompleter()
        self.textureModSizeCompleter_Hor4 = QCompleter()
        self.textureModAbundanceCompleter_Hor4 = QCompleter()
        self.structureShape1Completer_Hor4 = QCompleter()
        self.structureSize1Completer_Hor4 = QCompleter()
        self.structureGrade1Completer_Hor4 = QCompleter()
        self.structureRelationCompleter_Hor4 = QCompleter()
        self.structureShape2Completer_Hor4 = QCompleter()
        self.structureSize2Completer_Hor4 = QCompleter()
        self.structureGrade2Completer_Hor4 = QCompleter()
        self.consistenceMoistCompleter_Hor4 = QCompleter()
        self.consistenceStickinessCompleter_Hor4 = QCompleter()
        self.consistencePlasticityCompleter_Hor4 = QCompleter()
        self.mottleAbundanceCompleter_Hor4 = QCompleter()
        self.mottleSizeCompleter_Hor4 = QCompleter()
        self.mottleContrastCompleter_Hor4 = QCompleter()
        self.mottleShapeCompleter_Hor4 = QCompleter()
        self.mottleHueCompleter_Hor4 = QCompleter()
        self.mottleValueCompleter_Hor4 = QCompleter()
        self.mottleChromaCompleter_Hor4 = QCompleter()
        self.rootFineCompleter_Hor4 = QCompleter()
        self.rootMediumCompleter_Hor4 = QCompleter()
        self.rootCoarseCompleter_Hor4 = QCompleter()

        ### Horizon 5
        self.boundaryDistCompleter_Hor5 = QCompleter()
        self.boundaryTopoCompleter_Hor5 = QCompleter()
        self.SoilColorHue1Completer_Hor5 = QCompleter()
        self.SoilColorValue1Completer_Hor5 = QCompleter()
        self.SoilColorChroma1Completer_Hor5 = QCompleter()
        self.SoilColorHue2Completer_Hor5 = QCompleter()
        self.SoilColorValue2Completer_Hor5 = QCompleter()
        self.SoilColorChroma2Completer_Hor5 = QCompleter()
        self.SoilColorHue3Completer_Hor5 = QCompleter()
        self.SoilColorValue3Completer_Hor5 = QCompleter()
        self.SoilColorChroma3Completer_Hor5 = QCompleter()
        self.textureClassCompleter_Hor5 = QCompleter()
        self.textureSandCompleter_Hor5 = QCompleter()
        self.textureModSizeCompleter_Hor5 = QCompleter()
        self.textureModAbundanceCompleter_Hor5 = QCompleter()
        self.structureShape1Completer_Hor5 = QCompleter()
        self.structureSize1Completer_Hor5 = QCompleter()
        self.structureGrade1Completer_Hor5 = QCompleter()
        self.structureRelationCompleter_Hor5 = QCompleter()
        self.structureShape2Completer_Hor5 = QCompleter()
        self.structureSize2Completer_Hor5 = QCompleter()
        self.structureGrade2Completer_Hor5 = QCompleter()
        self.consistenceMoistCompleter_Hor5 = QCompleter()
        self.consistenceStickinessCompleter_Hor5 = QCompleter()
        self.consistencePlasticityCompleter_Hor5 = QCompleter()
        self.mottleAbundanceCompleter_Hor5 = QCompleter()
        self.mottleSizeCompleter_Hor5 = QCompleter()
        self.mottleContrastCompleter_Hor5 = QCompleter()
        self.mottleShapeCompleter_Hor5 = QCompleter()
        self.mottleHueCompleter_Hor5 = QCompleter()
        self.mottleValueCompleter_Hor5 = QCompleter()
        self.mottleChromaCompleter_Hor5 = QCompleter()
        self.rootFineCompleter_Hor5 = QCompleter()
        self.rootMediumCompleter_Hor5 = QCompleter()
        self.rootCoarseCompleter_Hor5 = QCompleter()

        ### Horizon 6
        self.boundaryDistCompleter_Hor6 = QCompleter()
        self.boundaryTopoCompleter_Hor6 = QCompleter()
        self.SoilColorHue1Completer_Hor6 = QCompleter()
        self.SoilColorValue1Completer_Hor6 = QCompleter()
        self.SoilColorChroma1Completer_Hor6 = QCompleter()
        self.SoilColorHue2Completer_Hor6 = QCompleter()
        self.SoilColorValue2Completer_Hor6 = QCompleter()
        self.SoilColorChroma2Completer_Hor6 = QCompleter()
        self.SoilColorHue3Completer_Hor6 = QCompleter()
        self.SoilColorValue3Completer_Hor6 = QCompleter()
        self.SoilColorChroma3Completer_Hor6 = QCompleter()
        self.textureClassCompleter_Hor6 = QCompleter()
        self.textureSandCompleter_Hor6 = QCompleter()
        self.textureModSizeCompleter_Hor6 = QCompleter()
        self.textureModAbundanceCompleter_Hor6 = QCompleter()
        self.structureShape1Completer_Hor6 = QCompleter()
        self.structureSize1Completer_Hor6 = QCompleter()
        self.structureGrade1Completer_Hor6 = QCompleter()
        self.structureRelationCompleter_Hor6 = QCompleter()
        self.structureShape2Completer_Hor6 = QCompleter()
        self.structureSize2Completer_Hor6 = QCompleter()
        self.structureGrade2Completer_Hor6 = QCompleter()
        self.consistenceMoistCompleter_Hor6 = QCompleter()
        self.consistenceStickinessCompleter_Hor6 = QCompleter()
        self.consistencePlasticityCompleter_Hor6 = QCompleter()
        self.mottleAbundanceCompleter_Hor6 = QCompleter()
        self.mottleSizeCompleter_Hor6 = QCompleter()
        self.mottleContrastCompleter_Hor6 = QCompleter()
        self.mottleShapeCompleter_Hor6 = QCompleter()
        self.mottleHueCompleter_Hor6 = QCompleter()
        self.mottleValueCompleter_Hor6 = QCompleter()
        self.mottleChromaCompleter_Hor6 = QCompleter()
        self.rootFineCompleter_Hor6 = QCompleter()
        self.rootMediumCompleter_Hor6 = QCompleter()
        self.rootCoarseCompleter_Hor6 = QCompleter()

    def setComboboxStyle(self):
        logging.info("Horizon Windows - Set combobox style")

        ### Horizon 1
        self.Boundary_Dist_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_1.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
        ### Horizon 2
        self.Boundary_Dist_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_2.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
        ### Horizon 3
        self.Boundary_Dist_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_3.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
        ### Horizon 4
        self.Boundary_Dist_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_4.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
        ### Horizon 5
        self.Boundary_Dist_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_5.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
        ### Horizon 6
        self.Boundary_Dist_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Boundary_Topo_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Texture_Class_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Sand_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Abundance_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Texture_Mod_Size_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Shape_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Relation_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Structure_Shape_2_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:160px}")
        self.Structure_Size_2_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Structure_Grade_2_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Moist_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Stickiness_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Consistence_Plasticity_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:140px}")
        self.Mottle_Abundance_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Size_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Contrast_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:110px}")
        self.Mottle_Shape_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Hue_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Value_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        # self.Mottle_Chroma_1_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:130px}")
        self.Root_Fine_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Medium_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        self.Root_Coarse_CB_Hor_6.setStyleSheet("QComboBox QAbstractItemView{border: 0px;color:black;min-width:120px}")
        
    def closeEvent(self, event):
        logging.info("Horizon Window closed")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    if not connection.createConnection():
        sys.exit(1)
    main = HorizonWindow()
    main.show()
    sys.exit(app.exec_())
