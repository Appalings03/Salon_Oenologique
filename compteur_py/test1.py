import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout

class MyWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.counter = 0

        # Créer un bouton et un label pour le compteur
        self.button = QPushButton('Incrémenter')
        self.label = QLabel(str(self.counter))

        # Connecter le bouton à une fonction de rappel
        self.button.clicked.connect(self.increment_counter)

        # Créer un layout vertical et ajouter le bouton et le label
        vbox = QVBoxLayout()
        vbox.addWidget(self.button)
        vbox.addWidget(self.label)

        # Appliquer le layout à la fenêtre
        self.setLayout(vbox)

    def increment_counter(self):
        self.counter += 1
        self.label.setText(str(self.counter))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec_())
