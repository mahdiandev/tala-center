import os
import sys
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFontDatabase, QFont
from ui import MainWindow


def main() -> None:
    """
    The main entry point of the Database application.

    Initializes the QApplication, displays the MainWindow, and starts the event loop.
    """
    os.environ['QT_QPA_PLATFORM'] = 'windows:darkmode=0'
    sys.argv += ['-platform', 'windows:darkmode=0']

    app = QApplication(sys.argv)

    fonts_dir = Path(__file__).resolve().parent / 'fonts'
    if fonts_dir.exists():
        for font_file in fonts_dir.glob('*'):
            if font_file.suffix.lower() in ('.ttf', '.otf', '.ttc'):
                QFontDatabase.addApplicationFont(str(font_file))

    default_font = QFont('Vazirmatn', 10)
    app.setFont(default_font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()