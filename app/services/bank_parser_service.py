from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import xlrd
from schemas.transaction import Transaction
from utils.shamsi import Shamsi
from utils.text import normalize_text, extract_digits


class BankParserService:
    """
    Identifies the bank type of financial statement files and extracts deposit transactions.
    Supports both legacy .xls and modern .xlsx formats.
    """

    def __init__(self) -> None:
        """
        Initializes the parser service.
        """
        pass

    def _is_numeric(self, val) -> bool:
        """
        Checks if a cell value represents a valid row number.
        """
        if val is None:
            return False

        if isinstance(val, (int, float)):
            return True

        val_str = str(val).strip()
        return val_str.isdigit()

    def _convert_xls_to_xlsx(self, xls_path: Path, xlsx_path: Path) -> None:
        """
        Reads a legacy .xls file, sanitizes data types, and converts it into .xlsx.

        Args:
            xls_path (Path): Path to the legacy .xls file.
            xlsx_path (Path): Destination path for the converted .xlsx file.
        """
        xls_book = xlrd.open_workbook(str(xls_path))
        xlsx_book = openpyxl.Workbook()

        first_sheet = True
        for sheet_name in xls_book.sheet_names():
            xls_sheet = xls_book.sheet_by_name(sheet_name)

            if first_sheet:
                xlsx_sheet = xlsx_book.active
                xlsx_sheet.title = sheet_name
                first_sheet = False
            else:
                xlsx_sheet = xlsx_book.create_sheet(title=sheet_name)

            for r in range(xls_sheet.nrows):
                for c in range(xls_sheet.ncols):
                    cell_value = xls_sheet.cell_value(r, c)

                    if isinstance(cell_value, float) and cell_value.is_integer():
                        cell_value = int(cell_value)

                    xlsx_sheet.cell(row=r + 1, column=c + 1, value=cell_value)

        xlsx_book.save(xlsx_path)
        xlsx_book.close()

    def _detect_bank_name(self, sheet: Worksheet) -> str | None:
        """
        Detects the bank name directly from the active Worksheet memory buffer.

        Args:
            sheet (Worksheet): The active worksheet in RAM.

        Returns:
            str | None: The bank name, or None if unrecognized.
        """
        h4_val = sheet.cell(row=4, column=8).value
        if h4_val and '119727908005' in str(h4_val):
            return 'ملی'

        e2_val = sheet.cell(row=2, column=5).value
        if e2_val and '2719-830-2221768-1' in str(e2_val):
            return 'ایران زمین'

        g15_val = sheet.cell(row=15, column=7).value
        if g15_val and '0279026634107' in str(g15_val):
            return 'تجارت'

        l2_val = sheet.cell(row=2, column=12).value
        if l2_val and '120703703000' in str(l2_val):
            return 'صادرات'

        return None

    def parse_excel(self, file_path: Path) -> list[Transaction]:
        """
        Parses the bank statement Excel file and returns normalized deposit transactions.

        Exactly resolves detection and parsing within a single file open-close cycle.

        Args:
            file_path (Path): Path to the bank statement Excel file.

        Returns:
            list[Transaction]: Extracted deposits.

        Raises:
            ValueError: If the bank type is not recognized or file is corrupted.
        """
        print(f'[BankParser] Starting parsing for file: {file_path.name}')
        is_xls = file_path.suffix.lower() == '.xls'
        target_path = file_path

        if is_xls:
            print(f'[BankParser] Conversion triggered for legacy XLS: {file_path.name}')
            target_path = file_path.with_suffix('.temp.xlsx')
            self._convert_xls_to_xlsx(file_path, target_path)

        try:
            wb = openpyxl.load_workbook(target_path, read_only=False, data_only=True)
            sheet: Worksheet = wb.active

            try:
                bank_name = self._detect_bank_name(sheet)
                print(f'[BankParser] Detection completed. Bank name: {bank_name}')

                if not bank_name:
                    raise ValueError(f'Bank statement type unrecognized for file: {file_path.name}')

                if bank_name == 'ملی':
                    transactions = self._parse_melli(sheet)
                elif bank_name == 'ایران زمین':
                    transactions = self._parse_iranzamin(sheet)
                elif bank_name == 'تجارت':
                    transactions = self._parse_tejarat(sheet)
                elif bank_name == 'صادرات':
                    transactions = self._parse_saderat(sheet)
                else:
                    transactions = []
            finally:
                wb.close()
        finally:
            if is_xls and target_path.exists():
                try:
                    target_path.unlink()
                except OSError:
                    pass

        return transactions

    def _parse_melli(self, sheet: Worksheet) -> list[Transaction]:
        """
        Extracts deposit rows from Melli Bank sheet.
        """
        transactions = []
        row_count = 0
        for row in sheet.iter_rows(min_row=1):
            row_count += 1
            if len(row) < 6:
                print(f'[Melli Debug] Row {row_count} skipped: columns count {len(row)} < 6')
                continue

            row_idx = row[0].value
            transaction_type = row[4].value
            if not self._is_numeric(row_idx) or not transaction_type:
                print(f'[Melli Debug] Row {row_count} skipped: row_idx={row_idx} (is_numeric={self._is_numeric(row_idx)}), transaction_type={transaction_type}')
                continue

            normalized_type = normalize_text(str(transaction_type))
            if normalized_type != 'واریز':
                print(f'[Melli Debug] Row {row_count} skipped: transaction_type normalized is "{normalized_type}" (expected "واریز")')
                continue

            try:
                amount_raw = row[5].value
                if amount_raw is None:
                    print(f'[Melli Debug] Row {row_count} skipped: amount cell is None')
                    continue

                amount_str = extract_digits(str(amount_raw))
                if not amount_str:
                    print(f'[Melli Debug] Row {row_count} skipped: no digits extracted from amount raw: {amount_raw}')
                    continue

                amount_val = int(amount_str)
                if amount_val <= 0:
                    print(f'[Melli Debug] Row {row_count} skipped: amount_val {amount_val} is not positive')
                    continue

                date_str = str(row[1].value or '').strip()
                time_str = str(row[2].value or '').strip()
                dt_str = f'{date_str} {time_str}'.strip()
                print(f'[Melli Debug] Row {row_count} parsing datetime string: "{dt_str}"')
                dt_obj = Shamsi.from_str(dt_str)

                transactions.append(
                    Transaction(
                        amount=amount_val,
                        date=dt_obj,
                        bank_name='ملی'
                    )
                )
                print(f'[Melli Debug] Row {row_count} successfully parsed: {amount_val} Rial on {dt_str}')
            except Exception as e:
                print(f'[Melli Debug] Row {row_count} failed with exception: {str(e)}')
                continue

        print(f'[Melli Debug] Finished parsing. Total rows scanned: {row_count}, Extracted: {len(transactions)}')
        return transactions

    def _parse_iranzamin(self, sheet: Worksheet) -> list[Transaction]:
        """
        Extracts deposit rows from Iran Zamin Bank sheet.
        """
        transactions = []
        row_count = 0
        for row in sheet.iter_rows(min_row=1):
            row_count += 1
            if len(row) < 5:
                print(f'[Iran Zamin Debug] Row {row_count} skipped: columns count {len(row)} < 5')
                continue

            row_idx = row[0].value
            if not self._is_numeric(row_idx):
                print(f'[Iran Zamin Debug] Row {row_count} skipped: row_idx={row_idx} is not numeric')
                continue

            try:
                amount_raw = row[4].value
                if amount_raw is None:
                    print(f'[Iran Zamin Debug] Row {row_count} skipped: amount cell is None')
                    continue

                amount_str = extract_digits(str(amount_raw))
                if not amount_str:
                    print(f'[Iran Zamin Debug] Row {row_count} skipped: no digits extracted from amount raw: {amount_raw}')
                    continue

                amount_val = int(amount_str)
                if amount_val <= 0:
                    print(f'[Iran Zamin Debug] Row {row_count} skipped: amount_val {amount_val} is not positive')
                    continue

                date_str = str(row[2].value or '').strip()
                time_str = str(row[3].value or '').strip()
                dt_str = f'{date_str} {time_str}'.strip()
                print(f'[Iran Zamin Debug] Row {row_count} parsing datetime string: "{dt_str}"')
                dt_obj = Shamsi.from_str(dt_str)

                transactions.append(
                    Transaction(
                        amount=amount_val,
                        date=dt_obj,
                        bank_name='ایران زمین'
                    )
                )
                print(f'[Iran Zamin Debug] Row {row_count} successfully parsed: {amount_val} Rial on {dt_str}')
            except Exception as e:
                print(f'[Iran Zamin Debug] Row {row_count} failed with exception: {str(e)}')
                continue

        print(f'[Iran Zamin Debug] Finished parsing. Total rows scanned: {row_count}, Extracted: {len(transactions)}')
        return transactions

    def _parse_tejarat(self, sheet: Worksheet) -> list[Transaction]:
        """
        Extracts deposit rows from Tejarat Bank sheet.
        """
        transactions = []
        row_count = 0
        for row in sheet.iter_rows(min_row=1):
            row_count += 1
            if len(row) < 16:
                print(f'[Tejarat Debug] Row {row_count} skipped: columns count {len(row)} < 16')
                continue

            row_idx = row[15].value
            if not self._is_numeric(row_idx):
                print(f'[Tejarat Debug] Row {row_count} skipped: row_idx={row_idx} is not numeric')
                continue

            try:
                amount_raw = row[10].value
                if amount_raw is None:
                    print(f'[Tejarat Debug] Row {row_count} skipped: amount cell is None')
                    continue

                amount_str = extract_digits(str(amount_raw))
                if not amount_str:
                    print(f'[Tejarat Debug] Row {row_count} skipped: no digits extracted from amount raw: {amount_raw}')
                    continue

                amount_val = int(amount_str)
                if amount_val <= 0:
                    print(f'[Tejarat Debug] Row {row_count} skipped: amount_val {amount_val} is not positive')
                    continue

                date_str = str(row[14].value or '').strip()
                time_str = str(row[13].value or '').strip()
                dt_str = f'{date_str} {time_str}'.strip()
                print(f'[Tejarat Debug] Row {row_count} parsing datetime string: "{dt_str}"')
                dt_obj = Shamsi.from_str(dt_str)

                transactions.append(
                    Transaction(
                        amount=amount_val,
                        date=dt_obj,
                        bank_name='تجارت'
                    )
                )
                print(f'[Tejarat Debug] Row {row_count} successfully parsed: {amount_val} Rial on {dt_str}')
            except Exception as e:
                print(f'[Tejarat Debug] Row {row_count} failed with exception: {str(e)}')
                continue

        print(f'[Tejarat Debug] Finished parsing. Total rows scanned: {row_count}, Extracted: {len(transactions)}')
        return transactions

    def _parse_saderat(self, sheet: Worksheet) -> list[Transaction]:
        """
        Extracts deposit rows from Saderat Bank sheet.
        """
        transactions = []
        row_count = 0
        for row in sheet.iter_rows(min_row=1):
            row_count += 1
            if len(row) < 9:
                print(f'[Saderat Debug] Row {row_count} skipped: columns count {len(row)} < 9')
                continue

            row_idx = row[0].value
            if not self._is_numeric(row_idx):
                print(f'[Saderat Debug] Row {row_count} skipped: row_idx={row_idx} is not numeric')
                continue

            try:
                amount_raw = row[8].value
                if amount_raw is None:
                    print(f'[Saderat Debug] Row {row_count} skipped: amount cell is None')
                    continue

                amount_str = extract_digits(str(amount_raw))
                if not amount_str:
                    print(f'[Saderat Debug] Row {row_count} skipped: no digits extracted from amount raw: {amount_raw}')
                    continue

                amount_val = int(amount_str)
                if amount_val <= 0:
                    print(f'[Saderat Debug] Row {row_count} skipped: amount_val {amount_val} is not positive')
                    continue

                date_str = str(row[1].value or '').strip()
                time_str = str(row[2].value or '').strip()
                dt_str = f'{date_str} {time_str}'.strip()
                print(f'[Saderat Debug] Row {row_count} parsing datetime string: "{dt_str}"')
                dt_obj = Shamsi.from_str(dt_str)

                transactions.append(
                    Transaction(
                        amount=amount_val,
                        date=dt_obj,
                        bank_name='صادرات'
                    )
                )
                print(f'[Saderat Debug] Row {row_count} successfully parsed: {amount_val} Rial on {dt_str}')
            except Exception as e:
                print(f'[Saderat Debug] Row {row_count} failed with exception: {str(e)}')
                continue

        print(f'[Saderat Debug] Finished parsing. Total rows scanned: {row_count}, Extracted: {len(transactions)}')
        return transactions