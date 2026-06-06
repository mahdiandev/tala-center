import os
import sys
import shutil
import subprocess
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QMessageBox, QFileDialog, QLabel, QLineEdit
)
from core.config import settings
from services.gold_price_service import GoldPriceService
from services.bank_parser_service import BankParserService
from services.gold_center_service import GoldCenterService

if sys.platform == 'win32':
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('goldcenter.app.v1')
    except Exception:
        pass


class ImportWorker(QThread):
    """
    Worker thread that handles updating gold prices, parsing statement files,
    and registering transactions without blocking the GUI.
    """
    status = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self) -> None:
        """
        Initializes the import worker.
        """
        super().__init__()

    def run(self) -> None:
        """
        Executes the import workflow in a background thread.
        """
        try:
            self.status.emit('در حال به روزرسانی تاریخچه طلا از سامانه TGJU...')
            gold_service = GoldPriceService()
            gold_service.sync_prices()

            self.status.emit('در حال بررسی فایل مرکز اصلی طلا...')
            gold_center_service = GoldCenterService(gold_service)

            self.status.emit('در حال بررسی فایل‌های صورت‌حساب ورودی...')
            input_folder = settings.input_folder
            if not input_folder.exists():
                self.error.emit('پوشه ورودی یافت نشد.')
                return

            excel_files = list(input_folder.glob('*.xls*'))
            if not excel_files:
                self.error.emit('هیچ فایل اکسلی در پوشه ورودی یافت نشد.')
                return

            parser = BankParserService()
            all_transactions = []

            for file in excel_files:
                self.status.emit(f'در حال پردازش فایل صورت‌حساب: {file.name}...')
                try:
                    transactions = parser.parse_excel(file)
                    all_transactions.extend(transactions)
                    self.status.emit(f'فایل {file.name} با موفقیت پردازش شد و {len(transactions)} تراکنش استخراج گردید.')
                except Exception as e:
                    self.status.emit(f'خطا در پردازش فایل {file.name}: {str(e)}')

            if not all_transactions:
                self.status.emit('هیچ تراکنش معتبری برای ثبت یافت نشد.')
                self.finished.emit({'added': 0, 'skipped': 0, 'bank_stats': {}})
                return

            self.status.emit('در حال بررسی موارد تکراری و ثبت تراکنش‌ها در مرکز اصلی طلا...')
            results = gold_center_service.add_transactions(all_transactions)
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))


class ExportWorker(QThread):
    """
    Worker thread that handles generating the export Excel sheet.
    """
    status = pyqtSignal(str)
    finished = pyqtSignal(int, str)
    error = pyqtSignal(str)

    def __init__(self) -> None:
        """
        Initializes the export worker.
        """
        super().__init__()

    def run(self) -> None:
        """
        Executes the export workflow in a background thread.
        """
        try:
            self.status.emit('در حال خروجی گرفتن...')
            gold_service = GoldPriceService()
            gold_center_service = GoldCenterService(gold_service)
            exported_count = gold_center_service.export_transactions()

            if exported_count == 0:
                self.status.emit('هیچ تراکنش جدیدی برای خروجی گرفتن یافت نشد.')
                self.finished.emit(0, '')
                return

            output_folder = settings.output_folder
            exported_files = list(output_folder.glob('tala-export-*.xlsx'))
            if not exported_files:
                self.error.emit('خطا در یافتن فایل خروجی گرفته شده.')
                return

            latest_file = max(exported_files, key=os.path.getmtime)
            self.status.emit(f'فایل خروجی با موفقیت در پوشه مربوطه ایجاد شد:\n{latest_file.name}')
            self.finished.emit(exported_count, str(latest_file))
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    """
    Main Application Window providing the interface to trigger bank import,
    export transactions, and dynamically select data directories.
    """

    def __init__(self) -> None:
        """
        Initializes the GUI widgets, layouts, and connects button actions.
        """
        super().__init__()
        self.setWindowTitle('مرکز طلا')
        self.resize(800, 600)
        self._init_ui()

    def _init_ui(self) -> None:
        """
        Constructs the layout, directory selection bar, action buttons, and standard console.
        """
        app_font = QFont('Segoe UI', 10)
        self.setFont(app_font)

        icon_path = Path(__file__).resolve().parent.parent / 'icon.ico'
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        folder_layout = QHBoxLayout()
        self.btn_select_folder = QPushButton('انتخاب پوشه', self)
        self.btn_select_folder.setFont(app_font)
        self.btn_select_folder.clicked.connect(self._handle_select_folder)

        self.txt_data_folder = QLineEdit(self)
        self.txt_data_folder.setReadOnly(True)
        self.txt_data_folder.setText(str(settings.data_folder))
        self.txt_data_folder.setFont(app_font)

        self.lbl_path = QLabel('مسیر پوشه داده‌ها:', self)
        self.lbl_path.setFont(app_font)

        folder_layout.addWidget(self.btn_select_folder)
        folder_layout.addWidget(self.txt_data_folder)
        folder_layout.addWidget(self.lbl_path)
        main_layout.addLayout(folder_layout)

        button_layout = QHBoxLayout()
        self.btn_import = QPushButton('ثبت تراکنش‌های بانکی', self)
        self.btn_import.setFixedHeight(45)
        self.btn_import.setFont(app_font)
        self.btn_import.clicked.connect(self._handle_import)

        self.btn_export = QPushButton('خروجی گرفتن از تراکنش‌ها', self)
        self.btn_export.setFixedHeight(45)
        self.btn_export.setFont(app_font)
        self.btn_export.clicked.connect(self._handle_export)

        button_layout.addWidget(self.btn_import)
        button_layout.addWidget(self.btn_export)
        main_layout.addLayout(button_layout)

        self.console = QTextEdit(self)
        self.console.setReadOnly(True)
        self.console.setFont(app_font)
        main_layout.addWidget(self.console)

        self.console.append('برنامه آماده کار است. یکی از عملیات‌های بالا را انتخاب کنید.')

    def _handle_select_folder(self) -> None:
        """
        Opens a directory selection dialog, updates active settings, and initializes subfolders.
        """
        selected_dir = QFileDialog.getExistingDirectory(
            self,
            'انتخاب پوشه داده‌ها',
            str(settings.data_folder)
        )
        if selected_dir:
            new_path = Path(selected_dir)
            settings.data_folder = new_path
            settings.save_config()
            self.txt_data_folder.setText(str(new_path))

            input_path = new_path / 'ورودی'
            output_path = new_path / 'خروجی'
            gold_center_file_path = new_path / 'مرکز اصلی طلا.xlsx'

            try:
                input_path.mkdir(parents=True, exist_ok=True)
                output_path.mkdir(parents=True, exist_ok=True)
            except OSError:
                pass

            if not gold_center_file_path.exists():
                template_file = Path(__file__).resolve().parent.parent / 'templates' / 'مرکز اصلی طلا.xlsx'
                if template_file.exists():
                    try:
                        shutil.copy2(template_file, gold_center_file_path)
                    except OSError:
                        pass

            self._update_console(f'مسیر پوشه داده‌ها با موفقیت تغییر یافت به:\n{str(new_path)}')

    def _handle_import(self) -> None:
        """
        Triggers the background ImportWorker thread.
        """
        self.btn_import.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.btn_select_folder.setEnabled(False)
        self.console.clear()

        self.import_worker = ImportWorker()
        self.import_worker.status.connect(self._update_console)
        self.import_worker.finished.connect(self._on_import_finished)
        self.import_worker.error.connect(self._on_worker_error)
        self.import_worker.start()

    def _on_import_finished(self, results: dict) -> None:
        """
        Updates the UI with a detailed import report.

        Args:
            results (dict): Dictionary containing addition statistics.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)

        if results.get('added', 0) == 0 and results.get('skipped', 0) == 0:
            return

        report = []
        report.append('\n' + '=' * 40)
        report.append('گزارش نهایی ثبت صورت‌حساب‌ها:')
        report.append(f'تعداد کل تراکنش‌های ثبت شده جدید: {results["added"]}')
        report.append(f'تعداد کل تراکنش‌های تکراری نادیده گرفته شده: {results["skipped"]}')
        report.append('-' * 40)
        report.append('آمار ثبت به تفکیک بانک:')

        for bank, count in results.get('bank_stats', {}).items():
            report.append(f'- {bank}: {count} تراکنش')

        report.append('=' * 40)
        self._update_console('\n'.join(report))

    def _handle_export(self) -> None:
        """
        Triggers the background ExportWorker thread.
        """
        self.btn_import.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.btn_select_folder.setEnabled(False)
        self.console.clear()

        self.export_worker = ExportWorker()
        self.export_worker.status.connect(self._update_console)
        self.export_worker.finished.connect(self._on_export_finished)
        self.export_worker.error.connect(self._on_worker_error)
        self.export_worker.start()

    def _on_export_finished(self, count: int, file_path: str) -> None:
        """
        Reports export success and opens the generated file.

        Args:
            count (int): Count of exported rows.
            file_path (str): File system path of the exported sheet.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)

        if count == 0:
            return

        self._update_console(f'\nعملیات خروجی گرفتن با موفقیت پایان یافت. تعداد {count} ردیف خروجی گرفته شد.')
        self._open_file(Path(file_path))

    def _on_worker_error(self, message: str) -> None:
        """
        Handles and displays worker errors.

        Args:
            message (str): The error message details.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)
        self._update_console(f'\nخطا در عملیات: {message}')
        QMessageBox.critical(self, 'خطا', message)

    def _update_console(self, text: str) -> None:
        """
        Appends status or report text to the read-only console window.

        Args:
            text (str): Log message to output.
        """
        self.console.append(text)

    def _open_file(self, path: Path) -> None:
        """
        Opens a file in the default OS file viewer.

        Args:
            path (Path): Path to the target file.
        """
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', str(path)])
            else:
                subprocess.call(['xdg-open', str(path)])
        except Exception as e:
            self._update_console(f'خطا در باز کردن خودکار فایل: {str(e)}')