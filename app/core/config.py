import os
import json
import shutil
from pathlib import Path
from pydantic import computed_field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """
    Configuration settings for the Database application.

    Loads configurations from Local AppData
    and resolves data directories relative to the data folder path.
    """
    data_folder: Path | None = None
    app_version: str = '1.1'
    
    def __init__(self, **values) -> None:
        """
        Initializes settings by reading the global config.json if it exists.
        """
        local_app_data = os.environ.get('LOCALAPPDATA')
        data_folder_path = None

        if local_app_data:
            config_dir = Path(local_app_data) / 'TalaCenter'
            config_file = config_dir / 'config.json'
            if config_file.exists():
                try:
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config_data = json.load(f)
                        if 'data_folder' in config_data:
                            data_folder_path = Path(config_data['data_folder'])
                except (json.JSONDecodeError, OSError):
                    pass

        super().__init__(data_folder=data_folder_path, **values)
        if self.data_folder:
            print(f'[Debug] Startup path check triggered. data_folder is set to: {self.data_folder}')
            self.initialize_directories()

    def initialize_directories(self) -> None:
        """
        Ensures input, output, and app data directories exist if data_folder is set.
        """
        if not self.data_folder:
            return

        print('[Debug] Ensuring required directories exist...')
        try:
            self.app_data_folder.mkdir(parents=True, exist_ok=True)
            self.data_folder.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            print(f'[Debug] Base directory creation failed: {e}')

        input_path = self.input_folder
        output_path = self.output_folder
        backup_path = self.backup_folder
        database_file_path = self.database_file

        if input_path and output_path and database_file_path:
            try:
                input_path.mkdir(parents=True, exist_ok=True)
                print(f'[Debug] Input folder verified: {input_path}')
                output_path.mkdir(parents=True, exist_ok=True)
                print(f'[Debug] Output folder verified: {output_path}')
                if backup_path:
                    backup_path.mkdir(parents=True, exist_ok=True)
                    print(f'[Debug] Backup folder verified: {backup_path}')
            except OSError as e:
                print(f'[Debug] Subfolder creation failed: {e}')

            if not database_file_path.exists():
                template_file = Path(__file__).resolve().parent.parent / 'templates' / 'database.xlsx'
                if template_file.exists():
                    try:
                        shutil.copy2(template_file, database_file_path)
                        print(f'[Debug] Database copied successfully to: {database_file_path}')
                    except OSError as e:
                        print(f'[Debug] Database template copy failed: {e}')
                else:
                    print(f'[Debug] Warning: Template database.xlsx not found at: {template_file}')

            shortcut_name = 'دیتابیس'
            shortcut_path = self.data_folder / shortcut_name
            print(f'[Debug] Database file verified. Creating shortcut at: {shortcut_path}')
            self.create_shortcut(database_file_path, shortcut_path)

    def create_shortcut(self, target: Path, shortcut_path: Path) -> None:
        """
        Creates a Windows shortcut (.lnk) file pointing to the target path.
        """
        print(f'[Debug] Attempting shortcut creation. target={target}, shortcut={shortcut_path}')
        
        temp_lnk = self.app_data_folder / 'temp_db.lnk'
        lnk_path = shortcut_path.with_suffix('.lnk')
        
        try:
            if temp_lnk.exists():
                temp_lnk.unlink()
            if lnk_path.exists():
                lnk_path.unlink()
        except OSError:
            pass

        created = False
        try:
            import win32com.client
            shell = win32com.client.Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(str(temp_lnk))
            shortcut.Targetpath = str(target)
            shortcut.save()
            print('[Debug] win32com shortcut creation succeeded on temporary path!')
            created = True
        except Exception as e:
            print(f'[Debug] win32com shortcut creation failed on temporary path: {e}. Trying PowerShell fallback...')
            try:
                import subprocess
                temp_lnk_str = str(temp_lnk).replace('\\', '/')
                target_str = str(target).replace('\\', '/')
                ps_command = (
                    f"$s = New-Object -ComObject WScript.Shell; "
                    f"$g = $s.CreateShortcut('{temp_lnk_str}'); "
                    f"$g.TargetPath = '{target_str}'; "
                    f"$g.Save()"
                )
                subprocess.run(
                    ['powershell', '-NoProfile', '-Command', ps_command],
                    capture_output=True,
                    text=True,
                    check=True
                )
                print('[Debug] PowerShell shortcut creation succeeded on temporary path!')
                created = True
            except Exception as pe:
                print(f'[Debug] PowerShell shortcut creation failed on temporary path: {pe}')

        if created and temp_lnk.exists():
            try:
                shutil.move(str(temp_lnk), str(lnk_path))
                print(f'[Debug] Shortcut successfully moved and renamed to: {lnk_path}')
            except Exception as me:
                print(f'[Debug] Moving shortcut to destination failed: {me}')

    def save_config(self) -> None:
        """
        Saves the current data_folder path back to the global configuration file.
        """
        if not self.data_folder:
            return

        config_dir = self.app_data_folder
        config_file = config_dir / 'config.json'

        try:
            config_dir.mkdir(parents=True, exist_ok=True)
            config_data = {'data_folder': str(self.data_folder)}
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
            self.initialize_directories()
        except OSError:
            pass

    @computed_field
    @property
    def app_data_folder(self) -> Path:
        """
        Path to the local AppData directory on Windows.
        """
        return Path(os.environ['LOCALAPPDATA']) / 'TalaCenter'

    @computed_field
    @property
    def input_folder(self) -> Path | None:
        """
        Path to the folder containing incoming bank statements.
        """
        if not self.data_folder:
            return None
        return self.data_folder / 'صورت حساب های بانکی'

    @computed_field
    @property
    def output_folder(self) -> Path | None:
        """
        Path to the folder where exported Excel files are saved.
        """
        if not self.data_folder:
            return None
        return self.data_folder / 'خروجی ها'

    @computed_field
    @property
    def backup_folder(self) -> Path | None:
        """
        Path to the folder where automated database backups are saved.
        """
        if not self.data_folder:
            return None
        return self.data_folder / 'بک آپ دیتابیس'

    @computed_field
    @property
    def database_file(self) -> Path:
        """
        Path to the main database file in the app data folder.
        """
        return self.app_data_folder / 'database.xlsx'

    @computed_field
    @property
    def gold_price_file(self) -> Path:
        """
        Path to the local JSON file storing gold price cache in the app data folder.
        """
        return self.app_data_folder / 'gold_price.json'

    model_config = SettingsConfigDict(
        case_sensitive=False,
        extra='ignore'
    )


settings = Settings()