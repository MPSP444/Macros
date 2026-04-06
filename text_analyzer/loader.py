import unicodedata


class TextLoader:
    ZERO_WIDTH = ['\u200b', '\u200c', '\u200d', '\ufeff', '\u00ad', '\u2060']
    UNUSUAL_SPACES = ['\u00a0', '\u2002', '\u2003', '\u2009', '\u202f']
    HOMOGLYPH_MAP = {
        # Cyrillic -> Latin lookalikes
        '\u0430': 'a',  # а -> a
        '\u0435': 'e',  # е -> e
        '\u043e': 'o',  # о -> o
        '\u0440': 'p',  # р -> p
        '\u0441': 'c',  # с -> c
        '\u0445': 'x',  # х -> x
        '\u0443': 'y',  # у -> y
        # Latin -> Cyrillic lookalikes
        'a': '\u0430',
        'e': '\u0435',
        'o': '\u043e',
        'p': '\u0440',
        'c': '\u0441',
        'x': '\u0445',
        'y': '\u0443',
    }

    def load(self, filepath: str) -> str:
        ext = filepath.rsplit('.', 1)[-1].lower()
        if ext == 'txt':
            return self._load_txt(filepath)
        elif ext == 'docx':
            return self._load_docx(filepath)
        elif ext == 'pdf':
            return self._load_pdf(filepath)
        else:
            raise ValueError(f"Unsupported file format: .{ext}")

    def _load_txt(self, filepath: str) -> str:
        for encoding in ('utf-8', 'cp1251', 'latin-1'):
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    return f.read()
            except (UnicodeDecodeError, LookupError):
                continue
        raise IOError(f"Cannot decode file: {filepath}")

    def _load_docx(self, filepath: str) -> str:
        from docx import Document
        doc = Document(filepath)
        return '\n'.join(p.text for p in doc.paragraphs)

    def _load_pdf(self, filepath: str) -> str:
        import fitz
        text_parts = []
        with fitz.open(filepath) as doc:
            for page in doc:
                text_parts.append(page.get_text())
        return '\n'.join(text_parts)

    def detect_hidden_symbols(self, text: str) -> dict:
        zero_width_found = []
        unusual_spaces_found = []
        homoglyphs_found = []

        for i, ch in enumerate(text):
            if ch in self.ZERO_WIDTH:
                zero_width_found.append({'pos': i, 'char': repr(ch), 'code': f'U+{ord(ch):04X}'})
            if ch in self.UNUSUAL_SPACES:
                unusual_spaces_found.append({'pos': i, 'char': repr(ch), 'code': f'U+{ord(ch):04X}'})

        # Detect homoglyphs: mixed-script words
        words = text.split()
        for word in words:
            has_cyr = any('\u0400' <= c <= '\u04FF' for c in word)
            has_lat = any('A' <= c <= 'Z' or 'a' <= c <= 'z' for c in word)
            if has_cyr and has_lat:
                homoglyphs_found.append({'word': word})

        return {
            'zero_width': zero_width_found,
            'homoglyphs': homoglyphs_found,
            'unusual_spaces': unusual_spaces_found,
        }

    def clean_hidden_symbols(self, text: str) -> str:
        for ch in self.ZERO_WIDTH:
            text = text.replace(ch, '')
        for ch in self.UNUSUAL_SPACES:
            text = text.replace(ch, ' ')
        return text
