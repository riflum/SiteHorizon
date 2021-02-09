
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtSql import QSqlDatabase, QSqlQuery


def createConnection():
    db = QSqlDatabase.addDatabase('QSQLITE')
    # Absolute path: is not working if the exe moved to another drive/folder -> fixed using relative path
    # db.setDatabaseName("I:\_Python\PyQT\SiteandHorizon\\New\Database\SiteHorizon.db")
    db.setDatabaseName("Database\SiteHorizon.db")
    if not db.open():
        QMessageBox.critical(None, "Cannot open database",
                "Unable to establish a database connection.\n"
                "This example needs SQLite support. Please read the Qt SQL "
                "driver documentation for information how to build it.\n\n"
                "Click Cancel to exit.",
                QMessageBox.Cancel)
        return False

    return True
