import os
import re
import json
import requests
from typing import List

SYNONYM_MAP = {
    'большой': ['крупный', 'значительный', 'обширный'],
    'маленький': ['небольшой', 'малый', 'незначительный'],
    'хороший': ['отличный', 'превосходный', 'качественный'],
    'плохой': ['неудовлетворительный', 'слабый', 'негативный'],
    'важный': ['существенный', 'значимый', 'ключевой'],
    'новый': ['современный', 'актуальный', 'инновационный'],
    'старый': ['устаревший', 'прежний', 'исторический'],
    'быстрый': ['оперативный', 'скоростной', 'стремительный'],
    'медленный': ['постепенный', 'неспешный', 'плавный'],
    'сложный': ['трудоёмкий', 'многогранный', 'непростой'],
    'простой': ['элементарный', 'базовый', 'доступный'],
    'получить': ['приобрести', 'достичь', 'обрести'],
    'использовать': ['применять', 'задействовать', 'эксплуатировать'],
    'показать': ['продемонстрировать', 'представить', 'отобразить'],
    'сделать': ['выполнить', 'осуществить', 'реализовать'],
    'разработать': ['создать', 'сформировать', 'построить'],
    'анализировать': ['исследовать', 'изучать', 'рассматривать'],
    'результат': ['итог', 'следствие', 'вывод'],
    'проблема': ['задача', 'вопрос', 'сложность'],
    'метод': ['способ', 'подход', 'техника'],
}


class TextModifier:
    def __init__(self):
        self._synonym_counters = {word: 0 for word in SYNONYM_MAP}

    def replace_synonyms(self, text: str) -> str:
        self._synonym_counters = {word: 0 for word in SYNONYM_MAP}

        def replacer(match):
            word = match.group(0)
            key = word.lower()
            if key in SYNONYM_MAP:
                synonyms = SYNONYM_MAP[key]
                idx = self._synonym_counters[key] % len(synonyms)
                self._synonym_counters[key] += 1
                replacement = synonyms[idx]
                # Preserve original case
                if word[0].isupper():
                    replacement = replacement[0].upper() + replacement[1:]
                return replacement
            return word

        pattern = re.compile(
            r'\b(' + '|'.join(re.escape(w) for w in SYNONYM_MAP.keys()) + r')\b',
            re.IGNORECASE
        )
        return pattern.sub(replacer, text)

    def diversify_sentences(self, text: str) -> str:
        sentences = re.split(r'(?<=[.!?])\s+', text.strip())
        result = []
        i = 0
        while i < len(sentences):
            s = sentences[i]
            word_count = len(s.split())

            if word_count > 30:
                # Split long sentence at comma
                parts = s.split(',')
                if len(parts) > 1:
                    mid = len(parts) // 2
                    first_part = ','.join(parts[:mid]).strip()
                    second_part = ','.join(parts[mid:]).strip()
                    if not first_part.endswith('.'):
                        first_part += '.'
                    if second_part and not second_part[0].isupper():
                        second_part = second_part[0].upper() + second_part[1:]
                    result.append(first_part)
                    result.append(second_part)
                else:
                    result.append(s)
                i += 1
            elif word_count < 5 and i + 1 < len(sentences):
                # Merge with next sentence
                next_s = sentences[i + 1]
                # Remove trailing punctuation from current, join
                merged = s.rstrip('.!?') + ', ' + next_s[0].lower() + next_s[1:]
                result.append(merged)
                i += 2
            else:
                result.append(s)
                i += 1

        return ' '.join(result)

    def paraphrase_gigachat(self, text: str, api_key: str = '') -> str:
        if not api_key:
            api_key = os.environ.get('GIGACHAT_API_KEY', '')

        if not api_key:
            return (
                "Для использования GigaChat API необходимо указать ключ.\n"
                "Установите переменную окружения GIGACHAT_API_KEY или введите ключ в поле API Key.\n\n"
                "Получить ключ можно на https://developers.sber.ru/portal/products/gigachat"
            )

        url = 'https://gigachat.devices.sberbank.ru/api/v1/chat/completions'
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
        }
        payload = {
            'model': 'GigaChat',
            'messages': [
                {
                    'role': 'user',
                    'content': (
                        'Перефразируй следующий текст, сохрани смысл, '
                        'но измени формулировки и структуру предложений:\n\n' + text
                    ),
                }
            ],
        }

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30, verify=False)
            response.raise_for_status()
            data = response.json()
            return data['choices'][0]['message']['content']
        except requests.exceptions.RequestException as e:
            return f"Ошибка при обращении к GigaChat API: {e}"
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            return f"Ошибка при обработке ответа GigaChat: {e}"
