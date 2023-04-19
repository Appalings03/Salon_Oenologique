from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit
import sys

class MainWindow(QWidget):
    def __init__(self, champ):
        super().__init__()
        self.champ = champ
        self.textes = {}
        self.initUI()

    def initUI(self):
        vbox = QVBoxLayout()
        for c in self.champ:
            label = QLabel(c)
            line_edit = QLineEdit()
            line_edit.textChanged.connect(lambda text, key=c: self.get_text(text, key)) # connect to get_text function
            vbox.addWidget(label)
            vbox.addWidget(line_edit)
        
        self.setLayout(vbox)
        self.show()

    def get_text(self, text, key):
        self.textes[key] = text # store text in dictionary with the corresponding key
        print(self.textes) # print the dictionary to check if the text is correctly stored

if __name__ == '__main__':
    app = QApplication(sys.argv)
    champ = ['Nom', 'Prénom', 'Âge']
    win = MainWindow(champ)
    sys.exit(app.exec())
