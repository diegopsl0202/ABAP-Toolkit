import sys

from PyQt5 import QtWidgets

from controller.screen import Screen

class Main:

    @staticmethod
    def execute():

        try:
            app = QtWidgets.QApplication(sys.argv)
            window = Screen(file =__file__)
            app.exec()

        except Exception as e:
            print(e)


if __name__ == '__main__':

    main = Main()
    main.execute()