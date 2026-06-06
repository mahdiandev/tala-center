import json
from pathlib import Path
from datetime import timedelta
import httpx
from core.config import settings
from utils.shamsi import Shamsi
from utils.text import extract_digits


class GoldPriceService:
    """
    Handles fetching, updating, and caching gold price history from the TGJU API,
    utilizing a centralized settings configuration and throwing exceptions on failures.
    """

    def __init__(self) -> None:
        """
        Initializes the service and resolves the cache path from global settings.
        """
        self.cache_path = settings.gold_history_file

    def _load_cache(self) -> dict[str, int]:
        """
        Loads cached gold prices from the local JSON file.

        Returns:
            dict[str, int]: A dictionary of Jalali date keys and price values.
        """
        if not self.cache_path.exists():
            return {}

        try:
            with open(self.cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}

    def _save_cache(self, data: dict[str, int]) -> None:
        """
        Saves the gold prices dictionary to the local JSON file.

        Args:
            data (dict[str, int]): The dictionary of gold prices to serialize.
        """
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except OSError:
            pass

    def fetch_prices_from_api(self) -> dict[str, int]:
        """
        Fetches gold price history from the TGJU API.

        Normalizes digits and filters values before returning.

        Returns:
            dict[str, int]: The fetched date-to-price dictionary.

        Raises:
            httpx.HTTPError: If the API request fails.
            ValueError: If the response data format is unexpected.
        """
        url = 'https://api.tgju.org/v1/market/indicator/summary-table-data/geram18?lang=fa&order_dir=asc&draw=1&start=0&length=300&search=&order_col=&order_dir=&from=&to=&convert_to_ad=1'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        response = httpx.get(url, headers=headers, timeout=15.0)
        response.raise_for_status()
        data_json = response.json()

        records = data_json.get('data')
        if not isinstance(records, list):
            raise ValueError('Invalid response format from TGJU API')

        api_data = {}
        for item in reversed(records):
            if not isinstance(item, list) or len(item) < 6:
                continue

            try:
                price_str = item[2]
                price_val = int(extract_digits(price_str)) - 500000
                date_key = item[-1]
                api_data[date_key] = price_val
            except (ValueError, IndexError):
                continue

        return api_data

    def fetch_today_price_from_api(self) -> int:
        """
        Fetches today's live gold price from the TGJU live market API.

        Normalizes the response, extracts digits, and returns the raw parsed value.

        Returns:
            int: Today's gold price, or 0 if the request or parsing fails.
        """
        url = 'https://api.tgju.org/v1/market/indicator/today-table-data/geram18?lang=fa&draw=1&start=0&length=30&search=&today_table_tolerance_open=1&today_table_tolerance_yesterday=1&today_table_tolerance_range=week'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        try:
            response = httpx.get(url, headers=headers, timeout=15.0)
            response.raise_for_status()
            data_json = response.json()

            records = data_json.get('data')
            if records and isinstance(records, list) and len(records) > 0:
                first_item = records[0]
                if isinstance(first_item, list) and len(first_item) > 0:
                    raw_price = first_item[0]
                    return int(extract_digits(raw_price))
        except (httpx.HTTPError, ValueError, IndexError):
            return 0

        return 0

    def sync_prices(self) -> int:
        """
        Synchronizes both the gold price history and today's live price with the local cache.

        Ensures today's date from Shamsi is paired with the live intraday gold price.
        Propagates explicit RuntimeError if the cache is outdated and sync fails.

        Returns:
            int: The total number of new or updated records in the cache.

        Raises:
            RuntimeError: If today's price is not cached and the network sync fails.
        """
        cache = self._load_cache()
        today_str = Shamsi.now().strftime('%Y/%m/%d')

        if today_str in cache:
            return 0

        try:
            api_data = self.fetch_prices_from_api()
        except Exception as e:
            raise RuntimeError('خطا در به روز رسانی تاریخچه طلا') from e

        try:
            today_price = self.fetch_today_price_from_api()
            if today_price > 0:
                today_str_live = Shamsi.now().strftime('%Y/%m/%d')
                api_data[today_str_live] = today_price
        except Exception:
            pass

        new_records_count = 0
        for date_key, price in api_data.items():
            if date_key not in cache or cache[date_key] != price:
                cache[date_key] = price
                new_records_count += 1

        if new_records_count > 0:
            self._save_cache(cache)

        return new_records_count

    def get_price(self, date_str: str) -> int:
        """
        Retrieves the gold price for a specific Jalali date.

        If the exact date is missing (e.g. holidays), loops backwards day-by-day 
        using timedelta subtraction to find the nearest preceding price. 
        Returns 0 if none found within 15 attempts.

        Args:
            date_str (str): The target Jalali date string.

        Returns:
            int: The gold price or 0 if missing.
        """
        normalized_date = date_str.replace('-', '/')
        cache = self._load_cache()

        if normalized_date in cache:
            return cache[normalized_date]

        try:
            shamsi_obj = Shamsi.from_str(normalized_date)
            attempt = 0
            max_attempts = 15
            current_str = normalized_date

            while current_str not in cache and attempt < max_attempts:
                shamsi_obj = shamsi_obj - timedelta(days=1)
                current_str = shamsi_obj.strftime('%Y/%m/%d')
                attempt += 1

            if current_str in cache:
                return cache[current_str]
        except Exception:
            pass

        return 0