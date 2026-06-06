import os
import json
import shutil
from pathlib import Path
from pydantic import computed_field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """
    Configuration settings for the Gold Center application.

    Loads configurations from ProgramData/TalaCenter/config.json and resolves
    data directories relative to the user's Documents folder by default.
    """
    data_folder: Path = Path.home() / 'Documents' / 'TalaCenterData'

    def __init__(self, **values) -> None:
        """
        Initializes settings by reading the global config.json in ProgramData.

        If the configuration file or folder does not exist, they are initialized
        with default values pointing to the Documents directory.
        Replicates templates and creates input/output directories if missing.
        """
        program_data_env = os.environ.get('ProgramData')
        if program_data_env:
            program_data_dir = Path(program_data_env)
        else:
            program_data_dir = Path('C:/ProgramData')

        config_dir = program_data_dir / 'TalaCenter'
        config_file = config_dir / 'config.json'
        
        default_data_folder = Path.home() / 'Documents' / 'TalaCenterData'
        data_folder_path = default_data_folder

        if config_file.exists():
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                    if 'data_folder' in config_data:
                        data_folder_path = Path(config_data['data_folder'])
            except (json.JSONDecodeError, OSError):
                pass
        else:
            try:
                config_dir.mkdir(parents=True, exist_ok=True)
                default_config = {'data_folder': str(default_data_folder)}
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, ensure_ascii=False, indent=4)
            except OSError:
                pass

        try:
            data_folder_path.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass

        input_path = data_folder_path / 'ورودی'
        output_path = data_folder_path / 'خروجی'
        gold_center_file_path = data_folder_path / 'مرکز اصلی طلا.xlsx'

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

        super().__init__(data_folder=data_folder_path, **values)

    def save_config(self) -> None:
        """
        Saves the current data_folder path back to the ProgramData configuration file.
        """
        program_data_env = os.environ.get('ProgramData')
        if program_data_env:
            program_data_dir = Path(program_data_env)
        else:
            program_data_dir = Path('C:/ProgramData')

        config_dir = program_data_dir / 'TalaCenter'
        config_file = config_dir / 'config.json'

        try:
            config_dir.mkdir(parents=True, exist_ok=True)
            config_data = {'data_folder': str(self.data_folder)}
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except OSError:
            pass

    @computed_field
    @property
    def input_folder(self) -> Path:
        """
        Path to the folder containing incoming bank statements.
        """
        return self.data_folder / 'ورودی'

    @computed_field
    @property
    def output_folder(self) -> Path:
        """
        Path to the folder where exported Excel files are saved.
        """
        return self.data_folder / 'خروجی'

    @computed_field
    @property
    def gold_center_file(self) -> Path:
        """
        Path to the main Gold Center Excel database file.
        """
        return self.data_folder / 'مرکز اصلی طلا.xlsx'

    @computed_field
    @property
    def gold_history_file(self) -> Path:
        """
        Path to the local JSON file storing gold price history.
        """
        return self.data_folder / 'تاریخچه طلا.json'

    model_config = SettingsConfigDict(
        case_sensitive=False,
        extra='ignore'
    )


settings = Settings()