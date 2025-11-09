"""
Configuration manager for ReviAI
"""
import configparser
from pathlib import Path
from typing import Dict, Any


class ConfigManager:
    """Manager for application configuration"""

    CONFIG_FILE = "config.ini"
    PROMPT_FILE = "prompt_template.txt"

    @classmethod
    def load_config(cls) -> Dict[str, Any]:
        """
        Load configuration from config.ini

        Returns:
            Dict containing configuration settings
        """
        config = configparser.ConfigParser()
        config_path = Path(cls.CONFIG_FILE)

        if not config_path.exists():
            raise FileNotFoundError(f"Configuration file not found: {cls.CONFIG_FILE}")

        config.read(config_path, encoding='utf-8')

        # Convert to dict for easier access
        config_dict = {
            'API': {
                'gemini_api_key': config.get('API', 'gemini_api_key'),
                'gemini_model': config.get('API', 'gemini_model', fallback='gemini-2.5-pro')
            },
            'Paths': {
                'default_output_dir': config.get('Paths', 'default_output_dir', fallback='./output')
            },
            'Settings': {
                'temperature': config.getint('Settings', 'temperature', fallback=0),
                'max_output_tokens': config.getint('Settings', 'max_output_tokens', fallback=8192),
                'max_retries': config.getint('Settings', 'max_retries', fallback=3)
            }
        }

        return config_dict

    @classmethod
    def save_config(cls, config_dict: Dict[str, Any]) -> None:
        """
        Save configuration to config.ini

        Args:
            config_dict: Configuration dictionary
        """
        config = configparser.ConfigParser()

        for section, options in config_dict.items():
            config[section] = options

        config_path = Path(cls.CONFIG_FILE)
        with open(config_path, 'w', encoding='utf-8') as f:
            config.write(f)

    @classmethod
    def validate_api_key(cls, api_key: str) -> bool:
        """
        Validate if API key is set

        Args:
            api_key: Gemini API key

        Returns:
            bool: True if valid, False otherwise
        """
        return api_key and api_key != "YOUR_API_KEY_HERE" and len(api_key) > 10

    @classmethod
    def get_prompt_template(cls) -> str:
        """
        Load prompt template from file

        Returns:
            str: Prompt template content
        """
        prompt_path = Path(cls.PROMPT_FILE)

        if not prompt_path.exists():
            raise FileNotFoundError(f"Prompt template not found: {cls.PROMPT_FILE}")

        with open(prompt_path, 'r', encoding='utf-8') as f:
            return f.read()

    @classmethod
    def save_prompt_template(cls, content: str) -> None:
        """
        Save prompt template to file

        Args:
            content: Prompt template content
        """
        prompt_path = Path(cls.PROMPT_FILE)

        with open(prompt_path, 'w', encoding='utf-8') as f:
            f.write(content)
