import sys
from PySide.QtCore import *
from PySide.QtGui import *
import csv

class WfmReport(QDialog):
    def __init__(self, data_list, header, parent=None):
        super(WfmReport, self).__init__(parent)

        #self.table_model = WFMReportModel(self, data_list, header)
        self.data_list = data_list
        self.header = header
        self.table_view = QTableView(self)
        self.table_model = QStandardItemModel(len(data_list), len(header), self.table_view)

        # set header
        for i in range(len(header)):
            self.table_model.setHeaderData(i, Qt.Horizontal, header[i])

        for row in range(len(data_list)):
            for col in range(len(header)):
                item = QStandardItem('%s' % (data_list[row][col]))
                item.setTextAlignment(Qt.AlignCenter)
                self.table_model.setItem(row, col, item)

        self.table_view.setModel(self.table_model)
        self.table_view.resizeColumnsToContents()
        btnSaveAs = QPushButton('Save As')
        btnExit = QPushButton('Exit')
        layoutButton = QHBoxLayout()
        #self.table_view.setSortingEnabled(True)

        #layoutButton.addSpacing(200)
        layoutButton.addStretch()
        layoutButton.addWidget(btnSaveAs)
        layoutButton.addWidget(btnExit)
        layoutButton.addStretch()
        #layoutButton.addItem(QSpacerItem(200,1), 0, 3)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table_view)
        layout.addLayout(layoutButton)
        self.setLayout(layout)

        btnSaveAs.clicked.connect(self.saveFile)
        btnExit.clicked.connect(self.exit)

        # align center all cell text
        #self.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def saveFile(self):
        filename,filter = QFileDialog.getSaveFileName(self, "Save Report",'',"CSV File (*.csv);; Text File (*.txt)")
        if filter.startswith("CSV"):
            outf = open(filename, "wb")
            writer = csv.writer(outf, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
            #write column headers
            writer.writerow(self.header)
            for row in self.data_list:
                writer.writerow(row)
        else:
            pass

        print "filename =", filename
        print "filter = ", filter

    def exit(self):
        self.close()
        self.reject()

    def setAlignment(self, alignment):
        r = 0
        for row in range(self.table_model.rowCount(self)):
            c = 0
            for col in range(self.table_model.columnCount(self)):
                self.table_view.item(r, c).setTextAlignment(alignment)
                c += 1
            r += 1


