import os
import sys
import re
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal, QMutex, QWaitCondition, Qt
from PyQt6.QtGui import QFont, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QMessageBox, QFileDialog, QLabel, QLineEdit,
    QDialogButtonBox, QDialog, QTableWidget, QTableWidgetItem, QHeaderView,
    QApplication
)
from core.config import settings
from services.gold_price_service import GoldPriceService
from services.bank_parser_service import BankParserService
from services.database_service import DatabaseService


def to_persian_numbers(text) -> str:
    """
    Translates standard English digits in a string into Persian digits.

    Args:
        text: The input text.

    Returns:
        str: The text with Persian digits.
    """
    if text is None:
        return ''
    translation_table = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')
    return str(text).translate(translation_table)


def to_english_numbers(text) -> str:
    """
    Translates Persian digits in a string into standard English digits.

    Args:
        text: The input text.

    Returns:
        str: The text with English digits.
    """
    if text is None:
        return ''
    translation_table = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
    return str(text).translate(translation_table)


def is_file_locked(file_path: Path) -> bool:
    """
    Checks if a file is currently open and locked by another process.

    Args:
        file_path (Path): Path to the target file.

    Returns:
        bool: True if locked, False otherwise.
    """
    if not file_path.exists():
        return False

    try:
        with open(file_path, 'r+'):
            pass
        return False
    except (PermissionError, OSError):
        return True


def get_resource_path(filename: str) -> Path:
    """
    Resolves the absolute path to a resource file, supporting both development
    and frozen PyInstaller environments via sys._MEIPASS fallback.

    Args:
        filename (str): The filename or relative path of the resource.

    Returns:
        Path: The resolved absolute path to the resource.
    """
    if hasattr(sys, '_MEIPASS'):
        base_path = Path(sys._MEIPASS)
        p_meipass_app = base_path / 'app' / filename
        if p_meipass_app.exists():
            return p_meipass_app
        p_meipass_root = base_path / filename
        if p_meipass_root.exists():
            return p_meipass_root

    current_dir = Path(__file__).resolve().parent
    p1 = current_dir / filename
    if p1.exists():
        return p1
    p2 = current_dir.parent / filename
    if p2.exists():
        return p2
    return p1


class FirstRunDialog(QDialog):
    """
    Welcomes the user on the first run and prompts them to select a data folder.
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle('خوش‌آمدید')
        self.setFixedSize(620, 160)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        
        icon_path = get_resource_path('icon.ico')
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
            
        self.selected_path = None
        self._init_ui()

    def _init_ui(self) -> None:
        app_font = QFont('Vazirmatn', 10)
        self.setFont(app_font)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        welcome_label = QLabel(self)
        welcome_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome_label.setText(
            '<div dir="rtl" style="text-align: center; line-height: 140%;">'
            '<span style="font-size: 11pt; font-weight: bold; color: #2e7d32;">به مرکز ثبت تراکنش‌های طلا خوش‌آمدید!</span><br>'
            '<span style="font-size: 10pt; font-weight: bold; color: #37474f;">'
            'لطفا آدرس پوشه داده های نرم افزار را انتخاب کنید.'
            '</span>'
            '</div>'
        )
        layout.addWidget(welcome_label)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        btn_style = (
            'QPushButton {'
            '    background-color: #f5f5f5;'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    padding: 8px 16px;'
            '    font-weight: bold;'
            '    color: #333333;'
            '}'
            'QPushButton:hover {'
            '    background-color: #e0e0e0;'
            '}'
            'QPushButton:focus {'
            '    border: 1px solid #b0b0b0;'
            '    outline: none;'
            '}'
        )

        btn_green_style = (
            'QPushButton {'
            '    background-color: #2e7d32;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 5px;'
            '    padding: 8px 16px;'
            '    font-weight: bold;'
            '}'
            'QPushButton:hover {'
            '    background-color: #1b5e20;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )

        self.btn_custom = QPushButton('انتخاب مسیر پوشه داده ها', self)
        self.btn_custom.setStyleSheet(btn_green_style)
        self.btn_custom.clicked.connect(self._on_custom_clicked)

        self.btn_cancel = QPushButton('خروج', self)
        self.btn_cancel.setStyleSheet(btn_style)
        self.btn_cancel.clicked.connect(self._on_cancel_clicked)

        btn_layout.addWidget(self.btn_custom)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)

    def _on_custom_clicked(self) -> None:
        selected_dir = QFileDialog.getExistingDirectory(
            self,
            'انتخاب پوشه داده‌ها',
            str(Path.home() / 'Documents')
        )
        if selected_dir:
            self.selected_path = Path(selected_dir)
            self.accept()

    def _on_cancel_clicked(self) -> None:
        self.reject()


class CumulativeSplitDialog(QDialog):
    """
    Dialog asking the user to manually split a cumulative POS transaction
    and validating that the sum of split amounts equals the total.
    """

    def __init__(self, parent, file_path: str, bank_name: str, total_amount: int, date_str: str, time_str: str, current_idx: int, total_count: int) -> None:
        super().__init__(parent)
        self.setWindowTitle('ثبت ریز تراکنش های تجمیعی')
        self.resize(500, 720)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        self.file_path = file_path
        self.total_amount = total_amount
        
        print(f'[Debug UI] CumulativeSplitDialog initialized: total_amount={total_amount} (type={type(total_amount)})')
        
        self.split_amounts = []
        self.current_idx = current_idx
        self.total_count = total_count
        file_name = Path(file_path).name
        self._init_ui(file_name, bank_name, date_str, time_str)

    def _init_ui(self, file_name: str, bank_name: str, date_str: str, time_str: str) -> None:
        app_font = QFont('Vazirmatn', 10)
        self.setFont(app_font)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        p_file_name = f'\u202A{file_name}\u202C'
        p_bank_name = to_persian_numbers(bank_name)
        p_total_amount = to_persian_numbers(f'{self.total_amount:,}')
        p_date = to_persian_numbers(date_str)
        p_time = to_persian_numbers(time_str)
        p_current_idx = to_persian_numbers(self.current_idx)
        p_total_count = to_persian_numbers(self.total_count)

        info_label = QLabel(self)
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setText(
            f'<div style="text-align: center; line-height: 150%;">'
            f'<span style="font-size: 14pt; font-weight: bold; color: #1565c0;">ثبت ریز تراکنش های تجمیعی</span><br>'
            f'<span style="font-size: 10pt; font-weight: bold; color: #757575;">(تراکنش {p_current_idx} از {p_total_count})</span><br>'
            f'<span style="font-size: 9pt; font-weight: bold; color: #37474f;">نام فایل: <span dir="ltr" style="display: inline-block;">{p_file_name}</span></span><br>'
            f'<span style="font-size: 9pt; font-weight: bold; color: #37474f;">مربوط به بانک {p_bank_name}</span><br>'
            f'<b style="font-size: 11pt; color: #2e7d32;">مبلغ تجمیعی: {p_total_amount} ریال</b><br>'
            f'<b style="color: #000000; font-size: 9pt;">زمان و تاریخ: {p_time} {p_date}</b>'
            f'</div>'
        )
        layout.addWidget(info_label)

        self.btn_open_file = QPushButton('📂 باز کردن فایل صورت حساب', self)
        self.btn_open_file.setFixedHeight(36)
        self.btn_open_file.setFont(app_font)
        self.btn_open_file.setStyleSheet(
            'QPushButton {'
            '    background-color: #f5f5f5;'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    font-weight: bold;'
            '    color: #333333;'
            '}'
            'QPushButton:hover {'
            '    background-color: #e0e0e0;'
            '}'
            'QPushButton:focus {'
            '    border: 1px solid #b0b0b0;'
            '    outline: none;'
            '}'
        )
        self.btn_open_file.clicked.connect(self._on_open_file_clicked)
        layout.addWidget(self.btn_open_file)

        instruction_label = QLabel(
            'ریز مبالغی که مربوط به این تراکنش تجمیعی هستند را وارد کنید:',
            self
        )
        instruction_label.setFont(QFont('Vazirmatn', 10, QFont.Weight.Bold))
        instruction_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(instruction_label)

        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)

        self.amount_input = QLineEdit(self)
        self.amount_input.setPlaceholderText('مبلغ به ریال')
        self.amount_input.setFixedHeight(42)
        self.amount_input.setFont(app_font)
        self.amount_input.setStyleSheet(
            'QLineEdit {'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    padding: 6px;'
            '    background-color: #fafafa;'
            '    font-size: 11pt;'
            '    font-weight: bold;'
            '}'
        )
        self.amount_input.textEdited.connect(self._on_amount_edited)
        self.amount_input.returnPressed.connect(self._on_add_clicked)

        self.btn_add = QPushButton('+', self)
        self.btn_add.setFixedSize(40, 42)
        self.btn_add.setFont(QFont('Vazirmatn', 22, QFont.Weight.Bold))
        self.btn_add.setStyleSheet(
            'QPushButton {'
            '    background-color: #2e7d32;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 5px;'
            '    font-weight: bold;'
            '}'
            'QPushButton:hover {'
            '    background-color: #1b5e20;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        self.btn_add.clicked.connect(self._on_add_clicked)

        input_layout.addWidget(self.btn_add)
        input_layout.addWidget(self.amount_input)
        layout.addLayout(input_layout)

        self.table_widget = QTableWidget(0, 3, self)
        self.table_widget.setFont(QFont('Vazirmatn', 11, QFont.Weight.Bold))
        self.table_widget.setHorizontalHeaderLabels(['ردیف', 'مبلغ', 'حذف'])
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_widget.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table_widget.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table_widget.setStyleSheet(
            'QTableWidget {'
            '    border: 1px solid #cccccc;'
            '    border-radius: 6px;'
            '    background-color: #ffffff;'
            '    font-size: 11pt;'
            '    font-weight: bold;'
            '    outline: none;'
            '}'
            'QTableWidget::item:selected {'
            '    background-color: #e0f2f1;'
            '    color: #000000;'
            '}'
            'QHeaderView::section {'
            '    background-color: #f5f5f5;'
            '    font-weight: bold;'
            '    border: 1px solid #cccccc;'
            '}'
        )
        layout.addWidget(self.table_widget)

        status_layout = QVBoxLayout()
        status_layout.setSpacing(5)

        self.lbl_sum = QLabel('مجموع مبالغ وارد شده: ۰ ریال', self)
        self.lbl_sum.setFont(QFont('Vazirmatn', 10, QFont.Weight.Bold))
        self.lbl_sum.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sum.setStyleSheet('color: #37474f;')
        self.lbl_sum.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.lbl_remaining = QLabel(f'مجموع مبالغ باقیمانده که باید وارد شود: {p_total_amount} ریال', self)
        self.lbl_remaining.setFont(QFont('Vazirmatn', 10, QFont.Weight.Bold))
        self.lbl_remaining.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_remaining.setStyleSheet('color: #c62828;')
        self.lbl_remaining.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        status_layout.addWidget(self.lbl_sum)
        status_layout.addWidget(self.lbl_remaining)
        layout.addLayout(status_layout)

        action_layout = QHBoxLayout()
        action_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_submit = QPushButton('همه مبالغ مربوط به این تراکنش را ثبت کردم', self)
        self.btn_submit.setFixedHeight(40)
        self.btn_submit.setFont(QFont('Vazirmatn', 10, QFont.Weight.Bold))
        self.btn_submit.setEnabled(False)
        self.btn_submit.setStyleSheet(
            'QPushButton {'
            '    background-color: #2e7d32;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 6px;'
            '    padding-left: 15px;'
            '    padding-right: 15px;'
            '}'
            'QPushButton:hover {'
            '    background-color: #1b5e20;'
            '}'
            'QPushButton:disabled {'
            '    background-color: #bdbdbd;'
            '    color: #f5f5f5;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        self.btn_submit.clicked.connect(self._on_submit_clicked)

        action_layout.addStretch()
        action_layout.addWidget(self.btn_submit)
        action_layout.addStretch()
        layout.addLayout(action_layout)

    def _on_amount_edited(self, text: str) -> None:
        cursor_pos = self.amount_input.cursorPosition()
        
        text_before_cursor = text[:cursor_pos]
        digits_before_cursor = sum(1 for c in text_before_cursor if c.isdigit() or c in '۰۱۲۳۴۵۶۷۸۹')

        clean_digits = ''.join(c for c in to_english_numbers(text) if c.isdigit())
        if clean_digits:
            val = int(clean_digits)
            formatted = to_persian_numbers(f'{val:,}')
            self.amount_input.setText(formatted)

            new_cursor_pos = 0
            digit_count = 0
            for i, char in enumerate(formatted):
                if char.isdigit() or char in '۰۱۲۳۴۵۶۷۸۹':
                    digit_count += 1
                new_cursor_pos = i + 1
                if digit_count == digits_before_cursor:
                    break

            self.amount_input.setCursorPosition(new_cursor_pos)
        else:
            self.amount_input.clear()

    def _on_add_clicked(self) -> None:
        text = self.amount_input.text()
        clean_digits = ''.join(c for c in to_english_numbers(text) if c.isdigit())
        if not clean_digits:
            return

        val = int(clean_digits)
        if val <= 0:
            return

        current_sum = sum(self.split_amounts)
        if current_sum + val > self.total_amount:
            QMessageBox.warning(self, 'خطا در مقدار', 'مجموع مبالغ وارد شده از کل مبلغ تجمیعی بیشتر می‌شود.')
            return

        self.split_amounts.append(val)
        
        row_idx = self.table_widget.rowCount()
        self.table_widget.insertRow(row_idx)

        item_num = QTableWidgetItem(to_persian_numbers(str(row_idx + 1)))
        item_num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        item_amount = QTableWidgetItem(to_persian_numbers(f'{val:,}'))
        item_amount.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        btn_del = QPushButton('🗑️', self)
        btn_del.setFixedSize(30, 24)
        btn_del.setStyleSheet(
            'QPushButton {'
            '    background-color: transparent;'
            '    border: none;'
            '}'
            'QPushButton:hover {'
            '    background-color: #ffebee;'
            '    border-radius: 3px;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        btn_del.clicked.connect(self._on_row_delete_clicked)

        self.table_widget.setItem(row_idx, 0, item_num)
        self.table_widget.setItem(row_idx, 1, item_amount)
        self.table_widget.setCellWidget(row_idx, 2, btn_del)

        self.amount_input.clear()
        self._update_status()

    def _on_row_delete_clicked(self) -> None:
        button = self.sender()
        if not button:
            return

        for r in range(self.table_widget.rowCount()):
            if self.table_widget.cellWidget(r, 2) == button:
                self.table_widget.removeRow(r)
                self.split_amounts.pop(r)
                break

        for r in range(self.table_widget.rowCount()):
            item_num = QTableWidgetItem(to_persian_numbers(str(r + 1)))
            item_num.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table_widget.setItem(r, 0, item_num)

        self._update_status()

    def _update_status(self) -> None:
        current_sum = sum(self.split_amounts)
        remaining = self.total_amount - current_sum

        self.lbl_sum.setText(to_persian_numbers(f'مجموع مبالغ وارد شده: {current_sum:,} ریال'))

        if remaining == 0:
            self.lbl_remaining.setText('مجموع مبالغ باقیمانده که باید وارد شود: ۰ ریال')
            self.lbl_remaining.setStyleSheet('color: #2e7d32; font-weight: bold;')
            self.btn_submit.setEnabled(True)
        else:
            remaining_persian = to_persian_numbers(f'{remaining:,}')
            self.lbl_remaining.setText(f'مجموع مبالغ باقیمانده که باید وارد شود: {remaining_persian} ریال')
            self.lbl_remaining.setStyleSheet('color: #c62828; font-weight: bold;')
            self.btn_submit.setEnabled(False)

    def _on_submit_clicked(self) -> None:
        try:
            print('[Debug UI] btn_submit clicked. Calculating sums...')
            current_sum = sum(self.split_amounts)
            print(f'[Debug UI] current_sum={current_sum}, total_amount={self.total_amount}')
            if current_sum != self.total_amount:
                print('[Debug UI] Warning: Sum does not match total amount.')
                return

            print(f'[Debug UI] Checking if file is locked: {self.file_path}')
            if is_file_locked(Path(self.file_path)):
                print('[Debug UI] Warning: File is locked.')
                error_text = 'فایل صورت حساب بانکی باز است. لطفاً آن را ببندید و دوباره بر روی دکمه ثبت کلیک کنید.'
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle('خطا در دسترسی به فایل')
                formatted_text = error_text.replace('\n', '<br>')
                html_content = (
                    f'<div dir="rtl" style="text-align: center; font-size: 10pt; font-weight: bold; '
                    f'line-height: 150%; min-width: 380px;">{formatted_text}</div>'
                )
                msg_box.setText(html_content)
                msg_box.setIcon(QMessageBox.Icon.Critical)
                button_box = msg_box.findChild(QDialogButtonBox)
                if button_box:
                    button_box.setCenterButtons(True)
                msg_box.exec()
                return

            print('[Debug UI] Accepting dialog...')
            self.accept()
        except Exception as e:
            import traceback
            print(f'[Debug UI Error] Exception in _on_submit_clicked: {e}')
            traceback.print_exc()

    def _on_open_file_clicked(self) -> None:
        try:
            os.startfile(self.file_path)
        except Exception as e:
            QMessageBox.warning(self, 'خطا', f'خطا در باز کردن فایل: {str(e)}')

    def get_split_amounts(self) -> list[int]:
        if self.result() == QDialog.DialogCode.Accepted:
            return self.split_amounts
        return []


class ImportWorker(QThread):
    """
    Worker thread that handles updating gold prices, parsing statement files,
    and registering transactions without blocking the GUI.
    """
    status = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    request_cumulative = pyqtSignal(str, str, object, str, str, int, int)
    request_summary = pyqtSignal(int)

    def __init__(self) -> None:
        """
        Initializes the import worker.
        """
        super().__init__()
        self.mutex = QMutex()
        self.wait_cond = QWaitCondition()
        self.split_result = None
        self.summary_approved = False
        self.total_cumulatives = 0
        self.current_cumulative_index = 0

    def handle_cumulative(self, file_path: str, bank_name: str, amount: int, date_str: str, time_str: str) -> list[int]:
        self.mutex.lock()
        self.split_result = None
        self.current_cumulative_index += 1
        self.request_cumulative.emit(
            file_path,
            bank_name,
            amount,
            date_str,
            time_str,
            self.current_cumulative_index,
            self.total_cumulatives
        )
        
        print('[Debug Worker] handle_cumulative: Thread is going to sleep, waiting for UI...')
        self.wait_cond.wait(self.mutex)
        
        result = self.split_result
        print(f'[Debug Worker] handle_cumulative: Thread woke up! Read split_result as: {result}')
        self.mutex.unlock()

        if not result:
            raise ValueError('فرایند ثبت تراکنش ها متوقف شد')

        return result

    def run(self) -> None:
        """
        Executes the import workflow in a background thread.
        """
        try:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                print('[Debug Worker] pythoncom CoInitialize called successfully.')
            except ImportError as ie:
                print(f'[Debug Worker] pythoncom import failed: {ie}')
            except Exception as ce:
                print(f'[Debug Worker] pythoncom CoInitialize failed: {ce}')

            if is_file_locked(settings.database_file):
                raise RuntimeError('فایل دیتابیس باز است. آن را ببندید و دوباره تلاش کنید.')

            self.status.emit('🔄 در حال به روزرسانی تاریخچه قیمت طلا...')
            gold_service = GoldPriceService()
            gold_service.sync_prices()

            self.status.emit('📁 در حال بررسی فایل دیتابیس...')
            database_service = DatabaseService(gold_service)

            self.status.emit('🔍 در حال بررسی فایل‌های صورت حساب بانکی...')
            input_folder = settings.input_folder
            if not input_folder or not input_folder.exists():
                self.error.emit('پوشه صورت حساب های بانکی یافت نشد.')
                return

            excel_files = [f for f in input_folder.glob('*.xls*') if not f.name.startswith('~$') and not f.name.startswith('.~')]
            if not excel_files:
                self.error.emit('هیچ فایل اکسلی در پوشه صورت حساب های بانکی یافت نشد.')
                return

            parser = BankParserService()
            cumulative_detections = []

            def scan_callback(f_path):
                def cb(bank_name, amount, date_str, time_str):
                    cumulative_detections.append({
                        'file_path': f_path,
                        'bank_name': bank_name,
                        'amount': amount,
                        'date_str': date_str,
                        'time_str': time_str
                    })
                    return []
                return cb

            total_files = len(excel_files)
            for idx, file in enumerate(excel_files, start=1):
                self.status.emit(f'🔍 در حال اسکن فایل صورت‌حساب ({idx} از {total_files}): {file.name}...')
                try:
                    parser.parse_excel(file, on_cumulative_found=scan_callback(str(file)))
                except Exception as e:
                    print(f'[Debug Worker] Scan failed for {file.name}: {e}')

            self.total_cumulatives = len(cumulative_detections)
            self.current_cumulative_index = 0

            if cumulative_detections:
                self.mutex.lock()
                self.summary_approved = False
                self.request_summary.emit(self.total_cumulatives)
                self.wait_cond.wait(self.mutex)
                approved = self.summary_approved
                self.mutex.unlock()

                if not approved:
                    raise ValueError('فرایند ثبت تراکنش ها متوقف شد')

            all_transactions = []

            for idx, file in enumerate(excel_files, start=1):
                self.status.emit(f'📥 در حال پردازش فایل صورت‌حساب ({idx} از {total_files}): {file.name}...')
                try:
                    def make_callback(f_path):
                        def callback(bank_name, amount, date_str, time_str):
                            return self.handle_cumulative(f_path, bank_name, amount, date_str, time_str)
                        return callback

                    transactions = parser.parse_excel(file, on_cumulative_found=make_callback(str(file)))
                    all_transactions.extend(transactions)
                    self.status.emit(f'✅ فایل {file.name} با موفقیت تحلیل و {len(transactions)} تراکنش استخراج شد.')
                except Exception as e:
                    self.status.emit(f'❌ خطا در پردازش فایل {file.name}: {str(e)}')
                    if 'متوقف شد' in str(e) or 'تراکنش تجمیعی' in str(e):
                        raise e

            if not all_transactions:
                self.status.emit('هیچ تراکنش معتبری برای ثبت یافت نشد.')
                self.finished.emit({'added': 0, 'skipped': 0, 'bank_stats': {}})
                return

            self.status.emit('📝 در حال بررسی موارد تکراری و ثبت تراکنش‌ها در فایل دیتابیس...')
            results = database_service.add_transactions(all_transactions)
            self.finished.emit(results)
        except Exception as e:
            import traceback
            print(f'[Debug Worker Error] Exception in ImportWorker.run: {e}')
            traceback.print_exc()
            self.error.emit(str(e))
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
                print('[Debug Worker] pythoncom CoUninitialize called successfully.')
            except Exception as ce:
                print(f'[Debug Worker] pythoncom CoUninitialize failed: {ce}')


class ExportWorker(QThread):
    """
    Worker thread that handles generating the export Excel sheet.
    """
    status = pyqtSignal(str)
    finished = pyqtSignal(int, list)
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
            if is_file_locked(settings.database_file):
                raise RuntimeError('فایل دیتابیس باز است. آن را ببندید و دوباره تلاش کنید.')

            self.status.emit('🔍 در حال استخراج ردیف‌های جدید از فایل دیتابیس...')
            gold_service = GoldPriceService()
            database_service = DatabaseService(gold_service)
            exported_count = database_service.export_transactions()

            if exported_count == 0:
                self.finished.emit(0, [])
                return

            output_folder = settings.output_folder
            if not output_folder:
                self.error.emit('پوشه خروجی معتبر نیست.')
                return

            exported_files = list(output_folder.glob('tala-export-*.xlsx'))
            if not exported_files:
                self.error.emit('خطا در یافتن فایل خروجی گرفته شده.')
                return

            latest_file = max(exported_files, key=os.path.getmtime)
            
            match = re.search(r'tala-export-(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})', latest_file.name)
            if not match:
                self.finished.emit(exported_count, [str(latest_file)])
                return

            base_timestamp = match.group(1)
            matching_files = []
            for file in exported_files:
                if base_timestamp in file.name:
                    matching_files.append(str(file))

            matching_files.sort()
            self.status.emit('📝 در حال ثبت ردیف‌ها در فایل خروجی...')
            self.finished.emit(exported_count, matching_files)
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    """
    Main Application Window providing the interface to trigger bank import,
    export transactions, and dynamically select data directories.
    """

    def __init__(self) -> None:
        """
        Initializes the GUI widgets, detects first-run configurations, and connects actions.
        """
        super().__init__()
        self.setWindowTitle('مرکز ثبت تراکنش‌های طلا')
        self.resize(800, 720)

        if settings.data_folder is None:
            self._handle_first_run()

        self._init_ui()

    def _handle_first_run(self) -> None:
        """
        Handles first-run welcome dialog and data folder selection.

        Exits the application if cancelled by the user.
        """
        dialog = FirstRunDialog(None)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_path:
            settings.data_folder = dialog.selected_path
            settings.save_config()
        else:
            sys.exit(0)

    def _init_ui(self) -> None:
        """
        Constructs the updated layout:
        1. Top: Centered icon.png and Bold Title Label.
        2. Middle: Large Import/Export buttons and read-only Console.
        3. Bottom: Directory selector and app version / author label.
        """
        app_font = QFont('Vazirmatn', 10)
        QApplication.setFont(app_font)
        self.setFont(app_font)

        icon_path = get_resource_path('icon.ico')
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        self.lbl_logo = QLabel(self)
        self.lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path = get_resource_path('icon.png')
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            scaled_pixmap = pixmap.scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.lbl_logo.setPixmap(scaled_pixmap)
        main_layout.addWidget(self.lbl_logo)

        self.lbl_title = QLabel('مرکز ثبت تراکنش‌های طلا', self)
        self.lbl_title.setFont(QFont('Vazirmatn', 16, QFont.Weight.Black))
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_title.setStyleSheet('color: #2e7d32;')
        main_layout.addWidget(self.lbl_title)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)

        self.btn_import = QPushButton('ثبت تراکنش‌های بانکی', self)
        self.btn_import.setFixedHeight(48)
        self.btn_import.setFont(QFont('Vazirmatn', 11, QFont.Weight.Bold))
        self.btn_import.setStyleSheet(
            'QPushButton {'
            '    background-color: #2e7d32;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 6px;'
            '    font-weight: bold;'
            '}'
            'QPushButton:hover {'
            '    background-color: #1b5e20;'
            '}'
            'QPushButton:disabled {'
            '    background-color: #bdbdbd;'
            '    color: #f5f5f5;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        self.btn_import.clicked.connect(self._handle_import)

        self.btn_export = QPushButton('خروجی گرفتن از تراکنش‌های طلا', self)
        self.btn_export.setFixedHeight(48)
        self.btn_export.setFont(QFont('Vazirmatn', 11, QFont.Weight.Bold))
        self.btn_export.setStyleSheet(
            'QPushButton {'
            '    background-color: #1565c0;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 6px;'
            '    font-weight: bold;'
            '}'
            'QPushButton:hover {'
            '    background-color: #0d47a1;'
            '}'
            'QPushButton:disabled {'
            '    background-color: #bdbdbd;'
            '    color: #f5f5f5;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        self.btn_export.clicked.connect(self._handle_export)

        button_layout.addWidget(self.btn_export)
        button_layout.addWidget(self.btn_import)
        main_layout.addLayout(button_layout)

        self.console = QTextEdit(self)
        self.console.setReadOnly(True)
        self.console.setFont(app_font)
        self.console.setStyleSheet(
            'QTextEdit {'
            '    border: 1px solid #cccccc;'
            '    border-radius: 6px;'
            '    background-color: #ffffff;'
            '    padding: 10px;'
            '}'
        )
        main_layout.addWidget(self.console)

        folder_layout = QHBoxLayout()
        folder_layout.setSpacing(10)

        self.btn_open_folder = QPushButton('📁 باز کردن پوشه داده‌ها', self)
        self.btn_open_folder.setFont(app_font)
        self.btn_open_folder.setStyleSheet(
            'QPushButton {'
            '    background-color: #00796b;'
            '    color: white;'
            '    border: none;'
            '    border-radius: 5px;'
            '    padding: 6px 14px;'
            '    font-weight: bold;'
            '}'
            'QPushButton:hover {'
            '    background-color: #004d40;'
            '}'
            'QPushButton:focus {'
            '    outline: none;'
            '}'
        )
        self.btn_open_folder.clicked.connect(self._handle_open_folder)
        
        self.btn_select_folder = QPushButton('تغییر مسیر پوشه داده‌ها', self)
        self.btn_select_folder.setFont(app_font)
        self.btn_select_folder.setStyleSheet(
            'QPushButton {'
            '    background-color: #f5f5f5;'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    padding: 6px 14px;'
            '    font-weight: bold;'
            '    color: #333333;'
            '}'
            'QPushButton:hover {'
            '    background-color: #e0e0e0;'
            '}'
            'QPushButton:focus {'
            '    border: 1px solid #b0b0b0;'
            '    outline: none;'
            '}'
        )
        self.btn_select_folder.clicked.connect(self._handle_select_folder)

        self.txt_data_folder = QLineEdit(self)
        self.txt_data_folder.setReadOnly(True)
        self.txt_data_folder.setText(str(settings.data_folder))
        self.txt_data_folder.setFont(app_font)
        self.txt_data_folder.setStyleSheet(
            'QLineEdit {'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    padding: 6px;'
            '    background-color: #fafafa;'
            '}'
        )

        self.lbl_path = QLabel('مسیر پوشه داده‌ها:', self)
        self.lbl_path.setFont(app_font)

        folder_layout.addWidget(self.btn_open_folder)
        folder_layout.addWidget(self.btn_select_folder)
        folder_layout.addWidget(self.txt_data_folder)
        folder_layout.addWidget(self.lbl_path)
        main_layout.addLayout(folder_layout)

        self.console.append('✨ لطفا یکی از دکمه های بالا را انتخاب کنید.')

        app_version = getattr(settings, 'app_version', '1.1')
        p_version = to_persian_numbers(app_version)
        self.lbl_version = QLabel(self)
        self.lbl_version.setFont(QFont('Vazirmatn', 9))
        self.lbl_version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_version.setOpenExternalLinks(True)
        self.lbl_version.setText(f'نسخه {p_version} - توسعه یافته توسط <a href="https://t.me/mahdiandev" style="color: #1565c0; text-decoration: none;">@mahdiandev</a>')
        main_layout.addWidget(self.lbl_version)

    def _show_message(self, title: str, text: str, is_error: bool = False) -> None:
        """
        Displays a styled, centered pop-up message box with centered buttons and larger text.

        Args:
            title (str): The window title of the message box.
            text (str): The text message to display.
            is_error (bool): If True, displays a critical error icon; otherwise, information icon.
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        
        formatted_text = text.replace('\n', '<br>')
        html_content = (
            f'<div dir="rtl" style="text-align: center; font-size: 10pt; font-weight: bold; '
            f'line-height: 150%; min-width: 380px;">{formatted_text}</div>'
        )
        msg_box.setText(html_content)

        if is_error:
            msg_box.setIcon(QMessageBox.Icon.Critical)
        else:
            msg_box.setIcon(QMessageBox.Icon.Information)

        button_box = msg_box.findChild(QDialogButtonBox)
        if button_box:
            button_box.setCenterButtons(True)

        msg_box.exec()

    def _handle_select_folder(self) -> None:
        """
        Opens a directory selection dialog and delegates path updates to settings configuration.
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
            self.console.append(f'📂 مسیر پوشه داده‌ها با موفقیت تغییر یافت به:\n{str(new_path)}')

    def _handle_open_folder(self) -> None:
        """
        Safely opens the currently configured data folder.
        """
        if settings.data_folder and settings.data_folder.exists():
            self._open_file(settings.data_folder)
        else:
            self._show_message('خطا', 'پوشه داده‌ها معتبر نیست یا هنوز ایجاد نشده است.', is_error=True)

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
        self.import_worker.request_cumulative.connect(self._on_cumulative_requested, Qt.ConnectionType.QueuedConnection)
        self.import_worker.request_summary.connect(self._on_summary_requested, Qt.ConnectionType.QueuedConnection)
        self.import_worker.start()

    def _on_summary_requested(self, count: int) -> None:
        p_count = to_persian_numbers(count)
        text_msg = f'تعداد {p_count} تراکنش تجمیعی در صورت‌حساب‌های بانکی وجود دارد.'
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('کشف تراکنش‌های تجمیعی')
        
        html_content = (
            f'<div style="text-align: center; font-size: 10pt; font-weight: bold; '
            f'line-height: 150%; min-width: 380px;">{text_msg}</div>'
        )
        msg_box.setText(html_content)
        msg_box.setIcon(QMessageBox.Icon.Information)
        
        btn_submit_splits = msg_box.addButton('ثبت ریز تراکنش های تجمیعی', QMessageBox.ButtonRole.AcceptRole)
        btn_submit_splits.setStyleSheet(
            'QPushButton {'
            '    background-color: #f5f5f5;'
            '    border: 1px solid #cccccc;'
            '    border-radius: 5px;'
            '    padding: 6px 14px;'
            '    font-weight: bold;'
            '    color: #333333;'
            '}'
            'QPushButton:hover {'
            '    background-color: #e0e0e0;'
            '}'
            'QPushButton:focus {'
            '    border: 1px solid #b0b0b0;'
            '    outline: none;'
            '}'
        )
        
        button_box = msg_box.findChild(QDialogButtonBox)
        if button_box:
            button_box.setCenterButtons(True)
            
        msg_box.exec()
        
        approved = (msg_box.clickedButton() == btn_submit_splits)
        
        self.import_worker.mutex.lock()
        self.import_worker.summary_approved = approved
        self.import_worker.wait_cond.wakeAll()
        self.import_worker.mutex.unlock()

    def _on_cumulative_requested(self, file_path: str, bank_name: str, amount: int, date_str: str, time_str: str, current_idx: int, total_count: int) -> None:
        try:
            print(f'[Debug UI] MainWindow received cumulative request: amount={amount} (type={type(amount)})')
            dialog = CumulativeSplitDialog(self, file_path, bank_name, amount, date_str, time_str, current_idx, total_count)
            print('[Debug UI] Executing CumulativeSplitDialog...')
            res = dialog.exec()
            print(f'[Debug UI] Dialog closed with result code: {res}')
            
            self.import_worker.mutex.lock()
            self.import_worker.split_result = dialog.get_split_amounts()
            print(f'[Debug UI] Returning split_result to worker: {self.import_worker.split_result}')
            self.import_worker.wait_cond.wakeAll()
            self.import_worker.mutex.unlock()
        except Exception as e:
            import traceback
            print(f'[Debug UI Error] Exception in _on_cumulative_requested: {e}')
            traceback.print_exc()
            self.import_worker.mutex.lock()
            self.import_worker.split_result = None
            self.import_worker.wait_cond.wakeAll()
            self.import_worker.mutex.unlock()

    def _on_import_finished(self, results: dict) -> None:
        """
        Displays import reports inside a success information message box.

        Args:
            results (dict): Dictionary containing addition statistics.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)

        added_count = results.get('added', 0)
        skipped_count = results.get('skipped', 0)

        if added_count > 0:
            p_added = to_persian_numbers(added_count)
            added_text_html = f'<span style="color: #2e7d32; font-weight: bold;">تعداد {p_added} تراکنش با موفقیت ثبت شد</span>'
            added_text_plain = f'تعداد {p_added} تراکنش با موفقیت ثبت شد'
        else:
            added_text_html = '<span style="color: #2e7d32; font-weight: bold;">هیچ تراکنشی ثبت نشد</span>'
            added_text_plain = 'هیچ تراکنشی ثبت نشد'

        report = []
        report.append('<br>')
        report.append(added_text_html)

        if skipped_count > 0:
            p_skipped = to_persian_numbers(skipped_count)
            report.append(f'<span style="color: #c62828;">تعداد تراکنش‌های تکراری: {p_skipped}</span>')

        if added_count > 0:
            report.append('<span style="color: #757575;">' + ('-' * 45) + '</span>')
            report.append('<b style="color: #1b5e20;">📈 آمار ثبت به تفکیک بانک:</b>')
            for bank, count in results.get('bank_stats', {}).items():
                p_bank = to_persian_numbers(bank)
                p_count = to_persian_numbers(count)
                report.append(f'<span style="color: #37474f;">- {p_bank}: {p_count} تراکنش</span>')

        report_html = '<br>'.join(report)
        self.console.append(report_html)

        plain_report = []
        plain_report.append(added_text_plain)

        if skipped_count > 0:
            p_skipped = to_persian_numbers(skipped_count)
            plain_report.append(f'تعداد تراکنش‌های تکراری: {p_skipped}')

        if added_count > 0:
            plain_report.append('-' * 45)
            plain_report.append('📈 آمار ثبت به تفکیک بانک:')
            for bank, count in results.get('bank_stats', {}).items():
                p_bank = to_persian_numbers(bank)
                p_count = to_persian_numbers(count)
                plain_report.append(f'- {p_bank}: {p_count} تراکنش')

        self._show_message('موفقیت در ثبت تراکنش‌ها', '\n'.join(plain_report))

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

    def _on_export_finished(self, count: int, file_paths: list) -> None:
        """
        Reports export success and opens the generated file.

        Args:
            count (int): Count of exported rows.
            file_paths (list): List of file system paths of the exported sheets.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)

        if count == 0:
            self.console.append('<br><b style="color: #c62828;">❌ هیچ تراکنش جدیدی برای خروجی گرفتن یافت نشد.</b>')
            self._show_message('هشدار', 'هیچ تراکنش جدیدی برای خروجی گرفتن یافت نشد.', is_error=True)
            return

        parts_count = len(file_paths)
        p_count = to_persian_numbers(count)
        p_parts = to_persian_numbers(parts_count)
        
        if parts_count > 1:
            report_text = f'تعداد {p_count} ردیف در {p_parts} پارت خروجی گرفته شد.'
            popup_text = f'🎉 عملیات خروجی با موفقیت پایان یافت.\n\nتعداد {p_count} ردیف در {p_parts} پارت خروجی گرفته شد.'
        else:
            report_text = f'تعداد {p_count} ردیف با موفقیت خروجی گرفته شد.'
            popup_text = f'🎉 عملیات خروجی با موفقیت پایان یافت.\n\nتعداد {p_count} ردیف با موفقیت خروجی گرفته شد.'

        report_html = (
            f'<br><span style="color: #2e7d32; font-weight: bold;">{report_text}</span>'
        )
        self.console.append(report_html)

        for path_str in file_paths:
            self._open_file(Path(path_str))

        self._show_message('موفقیت در خروجی', popup_text)

    def _on_worker_error(self, message: str) -> None:
        """
        Handles and displays worker errors.

        Args:
            message (str): The error message details.
        """
        self.btn_import.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_select_folder.setEnabled(True)
        p_message = to_persian_numbers(message)
        self.console.append(f'<br><b style="color: #c62828;">❌ خطا: {p_message}</b>')
        self._show_message('خطا', p_message, is_error=True)

    def _update_console(self, text: str) -> None:
        """
        Appends status or report text to the read-only console window.

        Args:
            text (str): Log message to output.
        """
        p_text = to_persian_numbers(text)
        self.console.append(p_text)

    def _open_file(self, path: Path) -> None:
        """
        Opens a file in the default Windows file viewer.

        Args:
            path (Path): Path to the target file.
        """
        try:
            os.startfile(path)
        except Exception as e:
            p_error = to_persian_numbers(str(e))
            self.console.append(f'❌ خطا در باز کردن خودکار فایل: {p_error}')