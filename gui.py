import main

import re
import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QGridLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QListView)
from PyQt5 import QtGui, QtCore

class MyApp(QWidget):
    input_file = ''
    styles = []
    items_model = QtGui.QStandardItemModel()

    def __init__(self):
        super().__init__()
        self.setUI()

    def setUI(self):
        self.setGeometry(300, 300, 300, 600)
        self.setWindowTitle('English Destroyer')
        window = QGridLayout()
        self.setLayout(window)

        self.file_name = QLineEdit(self.input_file)
        window.addWidget(QLabel('파일:'), 0, 0, 1, 1)
        window.addWidget(self.file_name, 0, 1, 1, -1)

        ipselectButton = QPushButton('파일 선택', self)
        ipselectButton.clicked.connect(self.showDialog)
        window.addWidget(ipselectButton, 1, 0, 1, -1)

        window.addWidget(QLabel('영어를 남겨둘 스타일'), 2, 0, 1, -1)

        self.list_view = QListView()
        self.list_view.setModel(self.items_model)
        window.addWidget(self.list_view, 3, 0, 6, -1)


        self.destroyer_button = QPushButton('영어 문단 제거하기', self)
        self.destroyer_button.clicked.connect(self.englishDestroyer)
        window.addWidget(self.destroyer_button, 10, 0, 1, -1)
        self.destroyer_button.setEnabled(False)

        self.show()


    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self, '.docx 파일을 선택해주세요', './', 'Word(*.docx)')
        self.input_file = fname[0]
        self.file_name.setText(self.input_file)
        if self.input_file:
            self.destroyer_button.setEnabled(True)
        
        try:
            self.styles = main.read_styles(self.input_file)
        except:
            QMessageBox.about(self, "에러 발생", "워드파일을 불러오지 못했습니다. 파일을 다시 선택해주세요.")
        else:
            self.items_model.clear()
            for style in self.styles:
                item = QtGui.QStandardItem(style)
                item.setCheckable(True)
                if re.search('코드', str(style)) != None:
                    item.setCheckState(2)
                self.items_model.appendRow(item)
            
    def englishDestroyer(self):
        safe_styles = []
        for row in range(self.items_model.rowCount()):
            item = self.items_model.item(row)
            if item.checkState() == 2:
                safe_styles.append(item.text())
        main.destroy_english(self.input_file, safe_styles)
        QMessageBox.about(self, "완료", "영어 문단을 모두 제거했습니다.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
