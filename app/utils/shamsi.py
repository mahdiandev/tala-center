from persiantools.jdatetime import JalaliDateTime
from zoneinfo import ZoneInfo
from datetime import datetime
from typing import Final
from .text import normalize_text

# --- Constants ---
TEHRAN_TIMEZONE: Final = ZoneInfo('Asia/Tehran')
UTC_TIMEZONE: Final = ZoneInfo('UTC')


class Shamsi(JalaliDateTime):
    '''
    An enhanced version of JalaliDateTime with smart formatting capabilities.
    '''

    def format(
        self,
        only_time: bool = False,
        only_date: bool = False,
        skip_milli_seconds: bool = False,
        in_words: bool = False,
    ) -> str:
        '''
        Formats the Shamsi datetime object into a custom string.
        
        Args:
            only_time (bool): If True, returns only the time part.
            only_date (bool): If True, returns only the date part.
            skip_milli_seconds (bool): If True, removes microseconds.
            in_words (bool): If True, uses Persian month names and locale formatting.
            
        Returns:
            str: The formatted datetime string.
        '''
        if in_words:
            if only_date:
                return self.strftime('%d %B %Y', locale='fa')
            
            if only_time:
                return self.strftime('%H:%M', locale='fa')
            
            return self.strftime('%c', locale='fa')

        instance_to_format = self
        if skip_milli_seconds:
            instance_to_format = self.replace(microsecond=0)

        if only_date:
            return instance_to_format.strftime('%Y-%m-%d')

        if only_time:
            if skip_milli_seconds:
                return instance_to_format.strftime('%H:%M:%S')
            return instance_to_format.strftime('%H:%M:%S.%f')[:-3]
        
        if skip_milli_seconds:
            return instance_to_format.strftime('%Y-%m-%d %H:%M:%S')
        
        return instance_to_format.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

    @classmethod
    def now(cls) -> 'Shamsi':
        '''
        Returns the current time in the Tehran timezone.
        '''
        return super().now(tz=TEHRAN_TIMEZONE)

    @classmethod
    def from_miladi(cls, dt: datetime) -> 'Shamsi':
        '''
        Converts a Miladi datetime object to a Shamsi instance.
        
        Ensures the input datetime is converted to Tehran timezone before 
        performing the Jalali conversion to maintain calendar accuracy.
        '''
        tehran_dt = dt.astimezone(TEHRAN_TIMEZONE)
        jalali_base = cls.to_jalali(tehran_dt)

        return cls(
            jalali_base.year,
            jalali_base.month,
            jalali_base.day,
            jalali_base.hour,
            jalali_base.minute,
            jalali_base.second,
            jalali_base.microsecond,
            tzinfo=TEHRAN_TIMEZONE
        )

    def to_miladi(self) -> datetime:
        '''
        Converts the Shamsi instance to a Miladi datetime object in UTC.
        '''
        naive_miladi = self.to_gregorian()
        aware_tehran = naive_miladi.replace(tzinfo=TEHRAN_TIMEZONE)
        
        return aware_tehran.astimezone(UTC_TIMEZONE)

    @classmethod
    def from_str(cls, dt_str: str) -> 'Shamsi':
        '''
        Parses a Jalali date or datetime string into a Shamsi instance.

        Supported formats (where separators can be '/' or '-'):
        - YYYY/MM/DD
        - YYYY/MM/DD HH:MM
        - YYYY/MM/DD HH:MM:SS

        Args:
            dt_str (str): The Jalali date or datetime string.

        Returns:
            Shamsi: The parsed Shamsi instance in Tehran timezone.

            Raises:
                ValueError: If the string format is invalid or cannot be parsed.
        '''

        cleaned = normalize_text(dt_str).strip()
        if not cleaned:
            raise ValueError('Input datetime string is empty')

        parts = cleaned.split()
        if len(parts) == 1:
            date_part = parts[0]
            time_part = '00:00:00'
        elif len(parts) == 2:
            date_part, time_part = parts
        else:
            raise ValueError(f'Invalid datetime format: {dt_str}')

        date_sep = '/' if '/' in date_part else '-'
        date_segments = date_part.split(date_sep)
        if len(date_segments) != 3:
            raise ValueError(f'Invalid date format: {date_part}')

        try:
            year = int(date_segments[0])
            month = int(date_segments[1])
            day = int(date_segments[2])
        except ValueError as e:
            raise ValueError(f'Date components must be numeric: {date_part}') from e

        time_segments = time_part.split(':')
        if len(time_segments) < 2 or len(time_segments) > 3:
            raise ValueError(f'Invalid time format: {time_part}')

        try:
            hour = int(time_segments[0])
            minute = int(time_segments[1])
            second = int(time_segments[2]) if len(time_segments) == 3 else 0
        except ValueError as e:
            raise ValueError(f'Time components must be numeric: {time_part}') from e

        try:
            return cls(
                year,
                month,
                day,
                hour,
                minute,
                second,
                tzinfo=TEHRAN_TIMEZONE
            )
        except Exception as e:
            raise ValueError(f'Invalid Jalali datetime components: {e}') from e
        
        
if __name__ == '__main__':
    date = Shamsi.from_str('1405/04/02 13:5:16')
    print(Shamsi.format(date))