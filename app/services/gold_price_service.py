import json
import re
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import httpx
from core.config import settings
from utils.shamsi import Shamsi
from utils.text import extract_digits


class GoldPriceService:
    """
    Handles fetching, synchronization, and caching gold price history from the TGJU API,
    utilizing a centralized settings configuration and throwing exceptions on failures.
    """

    def __init__(self) -> None:
        """
        Initializes the service and resolves the cache path from global settings.
        """
        self.cache_path = settings.gold_price_file

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

        Sorts the dictionary keys sequentially before serializing to preserve date order.

        Args:
            data (dict[str, int]): The dictionary of gold prices to serialize.
        """
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            sorted_data = {k: data[k] for k in sorted(data.keys())}
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(sorted_data, f, ensure_ascii=False, indent=4)
        except OSError:
            pass

    def fetch_prices_from_tgju(self) -> dict[str, int]:
        """
        Fetches gold price history from the TGJU API.

        Normalizes digits and filters values before returning.

        Returns:
            dict[str, int]: The fetched date-to-price dictionary.

        Raises:
            httpx.HTTPError: If the API request fails.
            ValueError: If the response data format is unexpected.
        """
        url = 'https://api.tgju.org/v1/market/indicator/summary-table-data/geram18?lang=fa&order_dir=asc&draw=1&start=0&length=200&search=&order_col=&order_dir=&from=&to=&convert_to_ad=1'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        print(f'[GoldPrice Debug] Fetching historical summary prices from TGJU API: {url}')
        response = httpx.get(url, headers=headers, timeout=15.0)
        print(f'[GoldPrice Debug] TGJU response status code: {response.status_code}')
        response.raise_for_status()
        data_json = response.json()

        records = data_json.get('data')
        if not isinstance(records, list):
            print('[GoldPrice Debug] Invalid or missing "data" key in TGJU response')
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

        print(f'[GoldPrice Debug] Successfully parsed {len(api_data)} records from TGJU API')
        return api_data

    def fetch_prices_from_milli_api(self) -> dict[str, int]:
        """
        Fetches live gold price history from the Milli Gold widget API.

        Converts millisecond timestamps into Jalali dates, groups multiple prices per day,
        calculates the daily average after 10:00 AM (or falls back to all records if none exist after 10:00 AM),
        and applies the Toman-to-Rial conversion.

        Returns:
            dict[str, int]: Jalali date-to-price mapping.
        """
        url = 'https://milli.gold/api/v1/public/milli-price/widget'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        print(f'[GoldPrice Debug] Fetching prices from Milli Gold API: {url}')
        response = httpx.get(url, headers=headers, timeout=15.0)
        print(f'[GoldPrice Debug] Milli Gold response status code: {response.status_code}')
        response.raise_for_status()
        data_json = response.json()

        data_node = data_json.get('data') or {}
        prices_node = data_node.get('prices') or {}
        
        monthly_list = prices_node.get('MONTHLY') or []
        weekly_list = prices_node.get('WEEKLY') or []
        daily_list = prices_node.get('DAILY') or []

        print(f'[GoldPrice Debug] Received monthly: {len(monthly_list)}, weekly: {len(weekly_list)}, daily: {len(daily_list)} records from Milli Gold')
        tehran_tz = ZoneInfo('Asia/Tehran')

        def process_milli_list(records_list) -> dict[str, float]:
            by_date = {}
            for item in records_list:
                val = item.get('value')
                ts_ms = item.get('timestamp')
                if val is None or ts_ms is None:
                    continue

                try:
                    ts_sec = ts_ms / 1000.0
                    dt_obj = datetime.fromtimestamp(ts_sec, tz=tehran_tz)
                    shamsi_date = Shamsi.from_miladi(dt_obj).strftime('%Y/%m/%d')

                    if shamsi_date not in by_date:
                        by_date[shamsi_date] = []
                    by_date[shamsi_date].append((dt_obj.hour, val))
                except Exception:
                    continue

            averages = {}
            for date_key, items_list in by_date.items():
                if not items_list:
                    continue

                filtered_vals = [val for hr, val in items_list if hr >= 10]

                if filtered_vals:
                    averages[date_key] = sum(filtered_vals) / len(filtered_vals)
                else:
                    all_vals = [val for hr, val in items_list]
                    averages[date_key] = sum(all_vals) / len(all_vals)

            return averages

        milli_data = {}

        monthly_averages = process_milli_list(monthly_list)
        for date_key, avg in monthly_averages.items():
            milli_data[date_key] = int(avg * 1000)

        weekly_averages = process_milli_list(weekly_list)
        for date_key, avg in weekly_averages.items():
            milli_data[date_key] = int(avg * 1000)

        daily_averages = process_milli_list(daily_list)
        for date_key, avg in daily_averages.items():
            milli_data[date_key] = int(avg * 1000)

        print(f'[GoldPrice Debug] Successfully calculated {len(milli_data)} daily average prices from Milli Gold')
        return milli_data

    def fetch_prices_from_taline_api(self) -> dict[str, int]:
        """
        Fetches gold price history from the Taline.ir website HTML source.

        Extracts the inline JavaScript 'const json = ...' chart variable using RegEx,
        converts Gregorian date keys to Jalali, and extracts the average price
        directly as Rials per gram without any subtraction.

        Returns:
            dict[str, int]: Jalali date-to-price mapping.
        """
        url = 'https://taline.ir/'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        print(f'[GoldPrice Debug] Loading HTML source from Taline.ir: {url}')
        response = httpx.get(url, headers=headers, timeout=15.0)
        print(f'[GoldPrice Debug] Taline.ir response status code: {response.status_code}')
        response.raise_for_status()
        html_content = response.text

        match = re.search(r'const\s+json\s*=\s*(\{"now":.*?history.*?\});', html_content, re.DOTALL)
        if not match:
            print('[GoldPrice Debug] Failed to find the inline JSON variable on Taline.ir via RegEx')
            return {}

        print('[GoldPrice Debug] Successfully matched inline JSON variable on Taline.ir')
        json_str = match.group(1)
        data = json.loads(json_str)
        history_node = data.get('history') or {}

        taline_data = {}
        tehran_tz = ZoneInfo('Asia/Tehran')

        for greg_date_str, values in history_node.items():
            avg_val = values.get('avg')
            if avg_val is None:
                continue

            try:
                dt_obj = datetime.strptime(greg_date_str, '%Y-%m-%d')
                dt_obj = dt_obj.replace(tzinfo=tehran_tz)
                shamsi_date = Shamsi.from_miladi(dt_obj).strftime('%Y/%m/%d')
                taline_data[shamsi_date] = int(avg_val)
            except Exception:
                continue

        print(f'[GoldPrice Debug] Successfully parsed {len(taline_data)} prices from Taline.ir')
        return taline_data

    def sync_prices(self) -> int:
        """
        Synchronizes cached gold prices.

        Fetches TGJU data first, then overwrites and supplements cache with daily average prices
        calculated from Milli Gold API, and finally supplements missing dates from Taline.ir.
        Sorts the final cache sequentially before saving.

        Returns:
            int: The total number of new or updated records in the cache.

        Raises:
            RuntimeError: If today's date is missing from the history cache after sync attempts.
        """
        cache = self._load_cache()
        today_str = Shamsi.now().strftime('%Y/%m/%d')
        print(f'[GoldPrice Debug] Starting price synchronization. Today JDateTime: {today_str}')

        if today_str in cache:
            print(f'[GoldPrice Debug] Today price ({today_str}) already exists in cache: {cache[today_str]}')
            return 0

        try:
            tgju_data = self.fetch_prices_from_tgju()
        except Exception as e:
            print(f'[GoldPrice Debug] TGJU sync failed or bypassed: {str(e)}')
            tgju_data = {}

        try:
            taline_data = self.fetch_prices_from_taline_api()
        except Exception as e:
            print(f'[GoldPrice Debug] Taline.ir sync failed or bypassed: {str(e)}')
            taline_data = {}

        try:
            milli_data = self.fetch_prices_from_milli_api()
        except Exception as e:
            print(f'[GoldPrice Debug] Milli Gold sync failed or bypassed: {str(e)}')
            milli_data = {}

        new_records_count = 0
        fetched_prices = {}

        for date_key, price in tgju_data.items():
            fetched_prices[date_key] = price

        for date_key, price in taline_data.items():
            fetched_prices[date_key] = price

        for date_key, price in milli_data.items():
            fetched_prices[date_key] = price

        print(f'[GoldPrice Debug] Blending {len(fetched_prices)} fetched prices into local cache database...')
        for date_key, price in fetched_prices.items():
            if date_key not in cache or cache[date_key] != price:
                cache[date_key] = price
                new_records_count += 1

        print(f'[GoldPrice Debug] Sync results: {new_records_count} records to write/update')
        if new_records_count > 0:
            self._save_cache(cache)
            print('[GoldPrice Debug] Cache file updated and sequentially sorted successfully')

        print(f'[GoldPrice Debug] Verifying if today date ({today_str}) is present in cache database...')
        if today_str not in cache:
            print(f'[GoldPrice Debug] Today date ({today_str}) is still missing from cache database after all attempts')
            raise RuntimeError('خطا در به روزرسانی تاریخچه طلا')

        print(f'[GoldPrice Debug] Today price is successfully secured: {cache[today_str]}')
        return new_records_count

    def get_price(self, date_str: str) -> int:
        """
        Retrieves the gold price for a specific Jalali date.

        If the exact date is missing (e.g. holidays), loops backwards day-by-day 
        using timedelta subtraction to find the nearest preceding price. 
        Raises RuntimeError if none found within 5 attempts.

        Args:
            date_str (str): The target Jalali date string.

        Returns:
            int: The gold price.

            Raises:
                RuntimeError: If the gold price for the date is missing from history after 5 attempts.
        """
        normalized_date = date_str.replace('-', '/')
        cache = self._load_cache()

        if normalized_date in cache:
            return cache[normalized_date]

        try:
            shamsi_obj = Shamsi.from_str(normalized_date)
            attempt = 0
            max_attempts = 5
            current_str = normalized_date

            while current_str not in cache and attempt < max_attempts:
                shamsi_obj = shamsi_obj - timedelta(days=1)
                current_str = shamsi_obj.strftime('%Y/%m/%d')
                attempt += 1

            if current_str in cache:
                return cache[current_str]
        except Exception:
            pass

        raise RuntimeError(f'قیمت طلای {date_str} در تاریخچه طلا موجود نیست')