import re
from persiantools import characters
from persian_tools import digits, separator


def extract_digits(text: str) -> str:
    """
    Extracts only digits from the normalized input text.

    Args:
        text (str): The input string to parse.

    Returns:
        str: A string containing only numeric digits.
    """
    return ''.join(re.findall(r'\d+', normalize_text(text)))


def extract_text(text: str) -> str:
    """
    Removes digits from the input text.

    Args:
        text (str): The input string to parse.

    Returns:
        str: The string with all digits removed.
    """
    return re.sub(r'\d', '', text)


def normalize_text(text: str | None) -> str:
    """
    Converts Arabic characters to Persian and digits to English.

    Ensures that empty or None inputs safely return an empty string to prevent
    downstream execution errors.

    Args:
        text (str | None): The raw input text.

    Returns:
        str: The normalized Persian/English text, or an empty string if input is None.
    """
    if not text:
        return ''

    normalized = characters.ar_to_fa(text)
    return digits.convert_to_en(normalized)


def normalize_amount(amount: str | int | None) -> int | None:
    """
    Normalizes and converts amount representation into an integer.

    Ensures consistent type conversion where strings containing non-digits
    and separators are safely converted to pure integers.

    Args:
        amount (str | int | None): The raw amount representation.

    Returns:
        int | None: The normalized integer value, or None if the input is None.
    """
    if amount is None:
        return None

    if isinstance(amount, int):
        return amount

    digits_str = extract_digits(amount)
    if not digits_str:
        return None

    return int(digits_str)


def format_amount(amount: str | int) -> str:
    """
    Formats the input amount with thousands separator.

    Args:
        amount (str | int): The raw amount representation.

    Returns:
        str: The formatted amount string with separators.
    """
    if isinstance(amount, str):
        amount_val = int(extract_digits(amount))
    else:
        amount_val = amount

    return f'{separator.add(amount_val)}'


def normalize_phone(phone: str | int | None) -> str:
    """
    Standardizes phone numbers by removing country codes, symbols, and leading zeros.

    Ensures comparison logic works regardless of whether the provider returns numbers
    as long integers or strings with prefixes.

    Args:
        phone (str | int | None): The raw phone number representation.

    Returns:
        str: The standardized 10-digit phone number string, or empty string if None.
    """
    if phone is None:
        return ''

    raw_digits = extract_digits(str(phone))
    if raw_digits.startswith('98') and len(raw_digits) > 10:
        raw_digits = raw_digits[2:]

    return raw_digits.lstrip('0')