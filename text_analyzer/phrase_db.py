import json
import os
from typing import List

DEFAULT_PHRASES = [
    "в данной работе рассматривается",
    "следует отметить, что",
    "таким образом, можно сделать вывод",
    "актуальность данной темы обусловлена",
    "в рамках данного исследования",
    "как показывает анализ",
    "необходимо подчеркнуть, что",
    "на основе вышеизложенного",
    "данная проблема является",
    "в соответствии с вышесказанным",
    "следует также отметить",
    "в результате проведённого анализа",
    "как было указано выше",
    "в ходе исследования было установлено",
    "принимая во внимание",
    "с точки зрения",
    "в заключение следует сказать",
    "на основании изложенного",
    "данная работа посвящена",
    "целью данной работы является",
    "предметом исследования является",
    "объектом исследования служит",
]


class PhraseDB:
    def __init__(self, path: str = 'phrases.json'):
        self.path = path
        self._phrases: List[str] = []
        self._load()

    def _load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, 'r', encoding='utf-8') as f:
                    self._phrases = json.load(f)
                return
            except (json.JSONDecodeError, IOError):
                pass
        self._phrases = list(DEFAULT_PHRASES)
        self._save()

    def _save(self):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(self._phrases, f, ensure_ascii=False, indent=2)

    def get_all(self) -> List[str]:
        return list(self._phrases)

    def add(self, phrase: str):
        phrase = phrase.strip()
        if phrase and phrase not in self._phrases:
            self._phrases.append(phrase)
            self._save()

    def remove(self, phrase: str):
        if phrase in self._phrases:
            self._phrases.remove(phrase)
            self._save()

    def update(self, old: str, new: str):
        new = new.strip()
        if old in self._phrases and new:
            idx = self._phrases.index(old)
            self._phrases[idx] = new
            self._save()

    def search(self, query: str) -> List[str]:
        q = query.lower()
        return [p for p in self._phrases if q in p.lower()]

    def reset_to_defaults(self):
        self._phrases = list(DEFAULT_PHRASES)
        self._save()

    def count(self) -> int:
        return len(self._phrases)
