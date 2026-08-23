import json
from pathlib import Path


SETTINGS_PATH = Path(__file__).parent / "settings.json"


def _load_saved_language() -> str:
    try:
        if SETTINGS_PATH.exists():
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
            return data.get("language", "ko")
    except Exception:
        pass
    return "ko"


def save_language(language: str) -> None:
    try:
        SETTINGS_PATH.write_text(
            json.dumps({"language": language}, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception:
        pass


class Localizer:
    def __init__(self, language: str = None):
        self.language = language or _load_saved_language()
        self.messages = {}
        self.load(self.language)

    def load(self, language: str) -> None:
        self.language = language
        path = Path(__file__).parent / "i18n" / f"{language}.json"

        if not path.exists():
            path = Path(__file__).parent / "i18n" / "en.json"

        with path.open("r", encoding="utf-8-sig") as file:
            self.messages = json.load(file)

    def t(self, key: str) -> str:
        return self.messages.get(key, key)


_localizer = Localizer()


def current_language() -> str:
    return _localizer.language


def set_language(language: str) -> None:
    _localizer.load(language)
    save_language(language)


def t(key: str) -> str:
    return _localizer.t(key)
