"""
语言管理模块。
负责加载和提供多语言翻译支持。
"""

import json
from pathlib import Path
from typing import Optional


class LanguageManager:
    """语言管理器"""

    def __init__(self, lang_file: Optional[str] = None, language: str = 'en'):
        if lang_file is None:
            self.lang_file = Path(__file__).parent / "language.json"
        else:
            self.lang_file = Path(lang_file)
        self.language = language
        self.translations: dict = {}
        self._load()

    def _load(self):
        """加载语言文件"""
        if self.lang_file.exists():
            with open(self.lang_file, 'r', encoding='utf-8') as f:
                self.translations = json.load(f)
        else:
            self.translations = {}

    def set_language(self, language: str):
        """切换当前语言"""
        self.language = language

    def get(self, key: str, **kwargs) -> str:
        """获取翻译文本，支持 {placeholder} 替换"""
        lang_dict = self.translations.get(self.language, {})
        text = lang_dict.get(key)
        if text is None:
            # 回退到英文
            text = self.translations.get('en', {}).get(key, key)
        if kwargs:
            text = text.format(**kwargs)
        return text

    def get_available_languages(self) -> list:
        """获取可用语言列表"""
        return list(self.translations.keys())

    def get_language_name(self, code: str) -> str:
        """获取语言的显示名称"""
        names = {
            'en': 'English',
            'zh': '中文',
        }
        return names.get(code, code)
