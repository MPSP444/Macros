import re
import math
from collections import Counter
from typing import List, Tuple, Dict

STOP_WORDS = {
    'и', 'в', 'во', 'не', 'что', 'он', 'на', 'я', 'с', 'со', 'как', 'а', 'то',
    'все', 'она', 'так', 'его', 'но', 'да', 'ты', 'к', 'у', 'же', 'вы', 'за',
    'бы', 'по', 'только', 'ее', 'мне', 'было', 'вот', 'от', 'меня', 'еще',
    'нет', 'о', 'из', 'ему', 'теперь', 'когда', 'даже', 'ну', 'вдруг', 'ли',
    'если', 'уже', 'или', 'ни', 'быть', 'был', 'него', 'до', 'вас', 'нибудь',
    'опять', 'уж', 'вам', 'ведь', 'там', 'потом', 'себя', 'ничего', 'ей', 'может',
    'они', 'тут', 'где', 'есть', 'надо', 'ней', 'для', 'мы', 'тебя', 'их', 'чем',
    'была', 'сам', 'чтоб', 'без', 'будто', 'чего', 'раз', 'тоже', 'себе', 'под',
    'будет', 'ж', 'тогда', 'кто', 'этот', 'того', 'потому', 'этого', 'какой',
    'совсем', 'ним', 'здесь', 'этом', 'один', 'почти', 'мой', 'тем', 'чтобы',
    'нее', 'кажется', 'сейчас', 'были', 'куда', 'зачем', 'всех', 'никогда',
    'можно', 'при', 'наконец', 'два', 'об', 'другой', 'хоть', 'после', 'над',
    'больше', 'тот', 'через', 'эти', 'нас', 'про', 'всего', 'них', 'какая',
    'много', 'разве', 'три', 'эту', 'моя', 'впрочем', 'хорошо', 'свою', 'этой',
    'перед', 'иногда', 'лучше', 'чуть', 'том', 'нельзя', 'такой', 'им', 'более',
    'всегда', 'конечно', 'всю', 'между',
    # English stop words
    'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
    'of', 'with', 'by', 'from', 'is', 'it', 'that', 'this', 'was', 'are',
    'be', 'as', 'had', 'he', 'she', 'they', 'we', 'you', 'i', 'not', 'no',
    'so', 'if', 'its', 'his', 'her', 'our', 'their', 'have', 'has', 'will',
    'do', 'did', 'been', 'can', 'could', 'would', 'should', 'may', 'might',
    'shall', 'into', 'than', 'then', 'there', 'when', 'where', 'which', 'who',
    'whom', 'all', 'each', 'every', 'both', 'few', 'more', 'most', 'other',
    'some', 'such', 'up', 'out', 'about', 'after', 'before', 'between', 'through',
}


def _tokenize(text: str) -> List[str]:
    return re.findall(r'\b[а-яёА-ЯЁa-zA-Z]+\b', text.lower())


def _split_sentences(text: str) -> List[str]:
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]


class TextAnalyzer:
    def __init__(self, text: str):
        self.text = text
        self._words = _tokenize(text)
        self._sentences = _split_sentences(text)

    def ttr(self) -> float:
        if not self._words:
            return 0.0
        return len(set(self._words)) / len(self._words)

    def burstiness(self) -> float:
        lengths = [len(_tokenize(s)) for s in self._sentences]
        if len(lengths) < 2:
            return 0.0
        mean = sum(lengths) / len(lengths)
        variance = sum((x - mean) ** 2 for x in lengths) / len(lengths)
        std = math.sqrt(variance)
        denom = std + mean
        if denom == 0:
            return 0.0
        return (std - mean) / denom

    def avg_sentence_length(self) -> float:
        if not self._sentences:
            return 0.0
        lengths = [len(_tokenize(s)) for s in self._sentences]
        return sum(lengths) / len(lengths)

    def top_words(self, n: int = 20) -> List[Tuple[str, int]]:
        filtered = [w for w in self._words if w not in STOP_WORDS and len(w) > 2]
        counter = Counter(filtered)
        return counter.most_common(n)

    def sentence_length_distribution(self) -> dict:
        lengths = [len(_tokenize(s)) for s in self._sentences]
        return {
            'lengths': lengths,
            'sentences': self._sentences,
        }

    def find_template_phrases(self, phrase_list: List[str]) -> List[dict]:
        found = []
        text_lower = self.text.lower()
        for phrase in phrase_list:
            pos = text_lower.find(phrase.lower())
            if pos != -1:
                found.append({'phrase': phrase, 'pos': pos})
        return found

    def score(self) -> dict:
        ttr_val = self.ttr()
        burst_val = self.burstiness()
        avg_len = self.avg_sentence_length()

        # TTR score: 0.7+ → 40 points
        ttr_score = 40 if ttr_val >= 0.7 else int(ttr_val / 0.7 * 40)

        # Burstiness score: -0.1..0.3 → 30 points
        if -0.1 <= burst_val <= 0.3:
            burst_score = 30
        else:
            distance = min(abs(burst_val - (-0.1)), abs(burst_val - 0.3))
            burst_score = max(0, int(30 - distance * 60))

        # Avg sentence length score: 12-20 words → 30 points
        if 12 <= avg_len <= 20:
            len_score = 30
        elif avg_len < 12:
            len_score = max(0, int(avg_len / 12 * 30))
        else:
            len_score = max(0, int(30 - (avg_len - 20) * 3))

        total = ttr_score + burst_score + len_score

        return {
            'total': min(100, total),
            'ttr_score': ttr_score,
            'burst_score': burst_score,
            'len_score': len_score,
            'ttr': round(ttr_val, 4),
            'burstiness': round(burst_val, 4),
            'avg_sentence_length': round(avg_len, 2),
            'word_count': len(self._words),
            'sentence_count': len(self._sentences),
            'unique_words': len(set(self._words)),
        }

    @staticmethod
    def compare(text1: str, text2: str) -> dict:
        a1 = TextAnalyzer(text1)
        a2 = TextAnalyzer(text2)
        s1 = a1.score()
        s2 = a2.score()

        words1 = set(_tokenize(text1))
        words2 = set(_tokenize(text2))
        intersection = words1 & words2
        union = words1 | words2
        jaccard = len(intersection) / len(union) if union else 0.0

        return {
            'text1': s1,
            'text2': s2,
            'jaccard': round(jaccard, 4),
        }
