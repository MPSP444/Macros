import os
import sys
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QFileDialog, QTableWidget,
    QTableWidgetItem, QStatusBar, QSplitter, QListWidget,
    QLineEdit, QComboBox, QMessageBox, QFrame, QGridLayout,
    QHeaderView, QSizePolicy, QInputDialog,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from loader import TextLoader
from analyzer import TextAnalyzer
from phrase_db import PhraseDB
from modifier import TextModifier
from exporter import ReportExporter

# ── Dark theme colours ──────────────────────────────────────────────────────
BG       = '#1e1e2e'
BG2      = '#181825'
BG3      = '#313244'
ACCENT   = '#89b4fa'
TEXT     = '#cdd6f4'
TEXT_DIM = '#6c7086'
SUCCESS  = '#a6e3a1'
WARNING  = '#f9e2af'
ERROR    = '#f38ba8'
CARD_BG  = '#24273a'

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG};
    color: {TEXT};
    font-family: "Segoe UI", "Liberation Sans", sans-serif;
    font-size: 13px;
}}
QTabWidget::pane {{
    border: 1px solid {BG3};
    background: {BG};
}}
QTabBar::tab {{
    background: {BG2};
    color: {TEXT_DIM};
    padding: 8px 20px;
    border: 1px solid {BG3};
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    min-width: 120px;
}}
QTabBar::tab:selected {{
    background: {BG};
    color: {ACCENT};
    border-bottom: 2px solid {ACCENT};
}}
QTabBar::tab:hover {{
    background: {BG3};
    color: {TEXT};
}}
QPushButton {{
    background-color: {BG3};
    color: {TEXT};
    border: 1px solid {ACCENT};
    border-radius: 5px;
    padding: 6px 16px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: {ACCENT};
    color: {BG};
}}
QPushButton:pressed {{
    background-color: #7aa2d4;
}}
QPushButton:disabled {{
    background-color: {BG3};
    color: {TEXT_DIM};
    border-color: {TEXT_DIM};
}}
QTextEdit, QLineEdit {{
    background-color: {BG2};
    color: {TEXT};
    border: 1px solid {BG3};
    border-radius: 4px;
    padding: 6px;
}}
QTextEdit:focus, QLineEdit:focus {{
    border: 1px solid {ACCENT};
}}
QTableWidget {{
    background-color: {BG2};
    color: {TEXT};
    border: 1px solid {BG3};
    gridline-color: {BG3};
    selection-background-color: {BG3};
}}
QTableWidget::item {{
    padding: 4px 8px;
}}
QHeaderView::section {{
    background-color: {BG3};
    color: {ACCENT};
    padding: 6px 8px;
    border: none;
    font-weight: bold;
}}
QListWidget {{
    background-color: {BG2};
    color: {TEXT};
    border: 1px solid {BG3};
    border-radius: 4px;
}}
QListWidget::item:selected {{
    background-color: {BG3};
    color: {ACCENT};
}}
QComboBox {{
    background-color: {BG2};
    color: {TEXT};
    border: 1px solid {BG3};
    border-radius: 4px;
    padding: 4px 8px;
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background-color: {BG2};
    color: {TEXT};
    selection-background-color: {BG3};
}}
QScrollBar:vertical {{
    background: {BG2};
    width: 10px;
}}
QScrollBar::handle:vertical {{
    background: {BG3};
    border-radius: 5px;
    min-height: 20px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QStatusBar {{
    background: {BG2};
    color: {TEXT_DIM};
    border-top: 1px solid {BG3};
}}
QSplitter::handle {{
    background: {BG3};
}}
"""


def _mpl_style(fig: Figure):
    fig.patch.set_facecolor(BG2)


def _ax_style(ax):
    ax.set_facecolor(BG2)
    ax.tick_params(colors=TEXT_DIM)
    ax.spines['bottom'].set_color(BG3)
    ax.spines['left'].set_color(BG3)
    ax.spines['top'].set_color(BG3)
    ax.spines['right'].set_color(BG3)
    ax.xaxis.label.set_color(TEXT_DIM)
    ax.yaxis.label.set_color(TEXT_DIM)
    ax.title.set_color(TEXT)


class MetricCard(QFrame):
    def __init__(self, title: str, value: str = '—', parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {BG3};
                border-radius: 8px;
                border: 1px solid #45475a;
                padding: 4px;
            }}
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)

        self.title_label = QLabel(title)
        self.title_label.setStyleSheet(f'color: {TEXT_DIM}; font-size: 11px;')
        self.title_label.setAlignment(Qt.AlignCenter)

        self.value_label = QLabel(value)
        self.value_label.setStyleSheet(f'color: {ACCENT}; font-size: 22px; font-weight: bold;')
        self.value_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value: str):
        self.value_label.setText(value)


class AnalysisWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, text: str, phrases: list):
        super().__init__()
        self.text = text
        self.phrases = phrases

    def run(self):
        try:
            loader = TextLoader()
            analyzer = TextAnalyzer(self.text)
            score = analyzer.score()
            top_words = analyzer.top_words(15)
            dist = analyzer.sentence_length_distribution()
            template_phrases = analyzer.find_template_phrases(self.phrases)
            manipulation = loader.detect_hidden_symbols(self.text)
            self.finished.emit({
                'score': score,
                'top_words': top_words,
                'dist': dist,
                'template_phrases': template_phrases,
                'manipulation': manipulation,
            })
        except Exception as e:
            self.error.emit(str(e))


# ════════════════════════════════════════════════════════════════════════════
#  Tab 1 – Analysis
# ════════════════════════════════════════════════════════════════════════════
class AnalysisTab(QWidget):
    text_analyzed = pyqtSignal(str)   # emitted so ModifyTab can pick it up

    def __init__(self, phrase_db: PhraseDB, status_bar: QStatusBar):
        super().__init__()
        self.phrase_db = phrase_db
        self.status_bar = status_bar
        self.loader = TextLoader()
        self._last_data = None
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # ── Top bar ──
        top = QHBoxLayout()
        self.btn_load = QPushButton('Загрузить файл')
        self.btn_load.setFixedHeight(32)
        self.btn_load.clicked.connect(self._load_file)

        self.btn_analyze = QPushButton('Анализировать')
        self.btn_analyze.setFixedHeight(32)
        self.btn_analyze.clicked.connect(self._analyze)

        self.btn_clean = QPushButton('Очистить текст')
        self.btn_clean.setFixedHeight(32)
        self.btn_clean.clicked.connect(self._clean_text)

        btn_export_pdf = QPushButton('Экспорт PDF')
        btn_export_pdf.setFixedHeight(32)
        btn_export_pdf.clicked.connect(self._export_pdf)

        btn_export_docx = QPushButton('Экспорт DOCX')
        btn_export_docx.setFixedHeight(32)
        btn_export_docx.clicked.connect(self._export_docx)

        top.addWidget(self.btn_load)
        top.addWidget(self.btn_analyze)
        top.addWidget(self.btn_clean)
        top.addStretch()
        top.addWidget(btn_export_pdf)
        top.addWidget(btn_export_docx)
        root.addLayout(top)

        # ── Split: left editor / right results ──
        splitter = QSplitter(Qt.Horizontal)

        # Left: text editor
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.addWidget(QLabel('Текст для анализа:'))
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText('Введите или загрузите текст...')
        lv.addWidget(self.text_edit)
        splitter.addWidget(left)

        # Right: results
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(6)

        # Metric cards
        cards_row = QHBoxLayout()
        self.card_score     = MetricCard('Score')
        self.card_ttr       = MetricCard('TTR')
        self.card_burst     = MetricCard('Burstiness')
        self.card_avg_len   = MetricCard('Ср. длина предл.')
        for card in (self.card_score, self.card_ttr, self.card_burst, self.card_avg_len):
            cards_row.addWidget(card)
        rv.addLayout(cards_row)

        # Details table
        rv.addWidget(QLabel('Детали:'))
        self.details_table = QTableWidget(3, 2)
        self.details_table.setHorizontalHeaderLabels(['Параметр', 'Значение'])
        self.details_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.details_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.details_table.setMaximumHeight(120)
        self.details_table.verticalHeader().setVisible(False)
        self._fill_details_table({})
        rv.addWidget(self.details_table)

        # Manipulation table
        rv.addWidget(QLabel('Детектор манипуляций:'))
        self.manip_table = QTableWidget(3, 2)
        self.manip_table.setHorizontalHeaderLabels(['Тип', 'Кол-во'])
        self.manip_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.manip_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.manip_table.setMaximumHeight(120)
        self.manip_table.verticalHeader().setVisible(False)
        self._fill_manip_table({})
        rv.addWidget(self.manip_table)

        # Charts
        charts_row = QHBoxLayout()

        self.fig1 = Figure(figsize=(4, 2.5), tight_layout=True)
        _mpl_style(self.fig1)
        self.canvas1 = FigureCanvas(self.fig1)
        self.canvas1.setMinimumHeight(160)

        self.fig2 = Figure(figsize=(4, 2.5), tight_layout=True)
        _mpl_style(self.fig2)
        self.canvas2 = FigureCanvas(self.fig2)
        self.canvas2.setMinimumHeight(160)

        charts_row.addWidget(self.canvas1)
        charts_row.addWidget(self.canvas2)
        rv.addLayout(charts_row)

        splitter.addWidget(right)
        splitter.setSizes([400, 700])
        root.addWidget(splitter)

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _fill_details_table(self, score: dict):
        data = [
            ('Слов', str(score.get('word_count', '—'))),
            ('Уникальных слов', str(score.get('unique_words', '—'))),
            ('Предложений', str(score.get('sentence_count', '—'))),
        ]
        for row, (k, v) in enumerate(data):
            self.details_table.setItem(row, 0, QTableWidgetItem(k))
            self.details_table.setItem(row, 1, QTableWidgetItem(v))

    def _fill_manip_table(self, manip: dict):
        data = [
            ('Нулевые пробелы', str(len(manip.get('zero_width', [])))),
            ('Омоглифы (слов)', str(len(manip.get('homoglyphs', [])))),
            ('Нестандартные пробелы', str(len(manip.get('unusual_spaces', [])))),
        ]
        for row, (k, v) in enumerate(data):
            self.details_table.setItem(row, 0, QTableWidgetItem(k)) if row == 0 else None
            self.manip_table.setItem(row, 0, QTableWidgetItem(k))
            item = QTableWidgetItem(v)
            if v != '0' and v != '—':
                item.setForeground(QColor(ERROR))
            self.manip_table.setItem(row, 1, item)

    def _plot_sentence_lengths(self, lengths: list):
        self.fig1.clear()
        ax = self.fig1.add_subplot(111)
        _ax_style(ax)
        if lengths:
            ax.hist(lengths, bins=min(20, len(set(lengths))),
                    color=ACCENT, edgecolor=BG2, alpha=0.85)
        ax.set_title('Длины предложений', fontsize=10, color=TEXT)
        ax.set_xlabel('Слов', fontsize=9)
        ax.set_ylabel('Частота', fontsize=9)
        self.canvas1.draw()

    def _plot_top_words(self, top_words: list):
        self.fig2.clear()
        ax = self.fig2.add_subplot(111)
        _ax_style(ax)
        if top_words:
            words = [w for w, _ in top_words[:15]]
            counts = [c for _, c in top_words[:15]]
            y_pos = range(len(words))
            ax.barh(list(y_pos), counts, color=ACCENT, alpha=0.85)
            ax.set_yticks(list(y_pos))
            ax.set_yticklabels(words, fontsize=8, color=TEXT)
            ax.invert_yaxis()
        ax.set_title('Топ-15 слов', fontsize=10, color=TEXT)
        ax.set_xlabel('Частота', fontsize=9)
        self.canvas2.draw()

    # ── Slots ─────────────────────────────────────────────────────────────────
    def _load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Открыть файл', '',
            'Текстовые файлы (*.txt *.docx *.pdf)'
        )
        if not path:
            return
        try:
            text = self.loader.load(path)
            self.text_edit.setPlainText(text)
            self.status_bar.showMessage(f'Загружен файл: {os.path.basename(path)}')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

    def _analyze(self):
        text = self.text_edit.toPlainText().strip()
        if not text:
            self.status_bar.showMessage('Введите текст для анализа')
            return
        self.btn_analyze.setEnabled(False)
        self.status_bar.showMessage('Анализ...')
        self._worker = AnalysisWorker(text, self.phrase_db.get_all())
        self._worker.finished.connect(self._on_analysis_done)
        self._worker.error.connect(self._on_analysis_error)
        self._worker.start()

    def _on_analysis_done(self, data: dict):
        self._last_data = data
        score = data['score']
        self.card_score.set_value(str(score['total']))
        self.card_ttr.set_value(f"{score['ttr']:.3f}")
        self.card_burst.set_value(f"{score['burstiness']:.3f}")
        self.card_avg_len.set_value(f"{score['avg_sentence_length']:.1f}")

        self._fill_details_table(score)
        self._fill_manip_table(data['manipulation'])
        self._plot_sentence_lengths(data['dist']['lengths'])
        self._plot_top_words(data['top_words'])

        self.btn_analyze.setEnabled(True)
        self.status_bar.showMessage(
            f"Готово. Score: {score['total']}  |  "
            f"Слов: {score['word_count']}  |  "
            f"Шаблонных фраз: {len(data['template_phrases'])}"
        )
        self.text_analyzed.emit(self.text_edit.toPlainText())

    def _on_analysis_error(self, msg: str):
        self.btn_analyze.setEnabled(True)
        self.status_bar.showMessage(f'Ошибка: {msg}')
        QMessageBox.critical(self, 'Ошибка анализа', msg)

    def _clean_text(self):
        text = self.text_edit.toPlainText()
        cleaned = self.loader.clean_hidden_symbols(text)
        self.text_edit.setPlainText(cleaned)
        self.status_bar.showMessage('Текст очищен от скрытых символов')

    def _export_pdf(self):
        if not self._last_data:
            self.status_bar.showMessage('Сначала выполните анализ')
            return
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить PDF', 'report.pdf', 'PDF (*.pdf)')
        if not path:
            return
        try:
            ReportExporter().export_pdf(path, self._last_data)
            self.status_bar.showMessage(f'PDF сохранён: {path}')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка экспорта', str(e))

    def _export_docx(self):
        if not self._last_data:
            self.status_bar.showMessage('Сначала выполните анализ')
            return
        path, _ = QFileDialog.getSaveFileName(self, 'Сохранить DOCX', 'report.docx', 'DOCX (*.docx)')
        if not path:
            return
        try:
            ReportExporter().export_docx(path, self._last_data)
            self.status_bar.showMessage(f'DOCX сохранён: {path}')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка экспорта', str(e))

    def get_text(self) -> str:
        return self.text_edit.toPlainText()


# ════════════════════════════════════════════════════════════════════════════
#  Tab 2 – Compare
# ════════════════════════════════════════════════════════════════════════════
class CompareTab(QWidget):
    def __init__(self, status_bar: QStatusBar):
        super().__init__()
        self.status_bar = status_bar
        self.loader = TextLoader()
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # Two text panels
        texts_row = QHBoxLayout()
        for attr, label in (('text1_edit', 'Текст 1'), ('text2_edit', 'Текст 2')):
            panel = QWidget()
            pv = QVBoxLayout(panel)
            pv.setContentsMargins(0, 0, 0, 0)
            pv.setSpacing(4)
            header = QHBoxLayout()
            header.addWidget(QLabel(label))
            btn = QPushButton('Загрузить файл')
            btn.setFixedHeight(28)
            btn.clicked.connect(lambda _, a=attr: self._load_file(a))
            header.addWidget(btn)
            pv.addLayout(header)
            te = QTextEdit()
            te.setPlaceholderText('Введите или загрузите текст...')
            setattr(self, attr, te)
            pv.addWidget(te)
            texts_row.addWidget(panel)
        root.addLayout(texts_row)

        # Compare button
        btn_compare = QPushButton('Сравнить')
        btn_compare.setFixedHeight(34)
        btn_compare.clicked.connect(self._compare)
        root.addWidget(btn_compare)

        # Results: table + chart
        results = QHBoxLayout()

        # Table
        self.cmp_table = QTableWidget(5, 3)
        self.cmp_table.setHorizontalHeaderLabels(['Метрика', 'Текст 1', 'Текст 2'])
        self.cmp_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.cmp_table.verticalHeader().setVisible(False)
        self.cmp_table.setMaximumWidth(480)
        results.addWidget(self.cmp_table)

        # Chart
        self.fig = Figure(figsize=(5, 3), tight_layout=True)
        _mpl_style(self.fig)
        self.canvas = FigureCanvas(self.fig)
        results.addWidget(self.canvas)

        root.addLayout(results)

    def _load_file(self, attr: str):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Открыть файл', '',
            'Текстовые файлы (*.txt *.docx *.pdf)'
        )
        if not path:
            return
        try:
            text = self.loader.load(path)
            getattr(self, attr).setPlainText(text)
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

    def _compare(self):
        t1 = self.text1_edit.toPlainText().strip()
        t2 = self.text2_edit.toPlainText().strip()
        if not t1 or not t2:
            self.status_bar.showMessage('Введите оба текста')
            return
        try:
            result = TextAnalyzer.compare(t1, t2)
            self._fill_table(result)
            self._plot(result)
            self.status_bar.showMessage(
                f"Jaccard similarity: {result['jaccard']}"
            )
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

    def _fill_table(self, r: dict):
        s1, s2 = r['text1'], r['text2']
        rows = [
            ('Score',                str(s1['total']),              str(s2['total'])),
            ('TTR',                  str(s1['ttr']),                str(s2['ttr'])),
            ('Burstiness',           str(s1['burstiness']),         str(s2['burstiness'])),
            ('Ср. длина предл.',     str(s1['avg_sentence_length']),str(s2['avg_sentence_length'])),
            ('Jaccard similarity',   str(r['jaccard']),             '—'),
        ]
        for row_idx, (m, v1, v2) in enumerate(rows):
            self.cmp_table.setItem(row_idx, 0, QTableWidgetItem(m))
            self.cmp_table.setItem(row_idx, 1, QTableWidgetItem(v1))
            self.cmp_table.setItem(row_idx, 2, QTableWidgetItem(v2))

    def _plot(self, r: dict):
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        _ax_style(ax)
        metrics = ['Score', 'TTR×100', 'Burstiness+1', 'Ср.длина']
        s1, s2 = r['text1'], r['text2']
        v1 = [s1['total'], s1['ttr'] * 100, (s1['burstiness'] + 1) * 50, s1['avg_sentence_length']]
        v2 = [s2['total'], s2['ttr'] * 100, (s2['burstiness'] + 1) * 50, s2['avg_sentence_length']]
        x = range(len(metrics))
        w = 0.35
        ax.bar([i - w / 2 for i in x], v1, w, label='Текст 1', color=ACCENT, alpha=0.85)
        ax.bar([i + w / 2 for i in x], v2, w, label='Текст 2', color=SUCCESS, alpha=0.85)
        ax.set_xticks(list(x))
        ax.set_xticklabels(metrics, fontsize=9, color=TEXT_DIM)
        ax.legend(facecolor=BG3, labelcolor=TEXT, fontsize=9)
        ax.set_title('Сравнение метрик', fontsize=11, color=TEXT)
        self.canvas.draw()


# ════════════════════════════════════════════════════════════════════════════
#  Tab 3 – Modify
# ════════════════════════════════════════════════════════════════════════════
class ModifyTab(QWidget):
    def __init__(self, status_bar: QStatusBar):
        super().__init__()
        self.status_bar = status_bar
        self.modifier = TextModifier()
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # Controls row
        ctrl = QHBoxLayout()
        ctrl.addWidget(QLabel('Режим:'))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems([
            'Синонимы (локально)',
            'GigaChat API',
            'Разнообразие предложений',
        ])
        self.mode_combo.setFixedWidth(220)
        ctrl.addWidget(self.mode_combo)

        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText('GigaChat API Key (опционально)')
        self.api_key_input.setEchoMode(QLineEdit.Password)
        ctrl.addWidget(self.api_key_input)

        btn_apply = QPushButton('Применить')
        btn_apply.clicked.connect(self._apply)
        ctrl.addWidget(btn_apply)

        btn_copy = QPushButton('Копировать')
        btn_copy.clicked.connect(self._copy_result)
        ctrl.addWidget(btn_copy)

        btn_transfer = QPushButton('Из вкладки Анализ')
        btn_transfer.clicked.connect(self._request_transfer)
        ctrl.addWidget(btn_transfer)

        root.addLayout(ctrl)

        # Editors
        editors = QSplitter(Qt.Horizontal)

        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.addWidget(QLabel('Исходный текст:'))
        self.src_edit = QTextEdit()
        lv.addWidget(self.src_edit)
        editors.addWidget(left)

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.addWidget(QLabel('Результат:'))
        self.dst_edit = QTextEdit()
        self.dst_edit.setReadOnly(True)
        rv.addWidget(self.dst_edit)
        editors.addWidget(right)

        root.addWidget(editors)

        self._transfer_callback = None

    def set_transfer_callback(self, cb):
        self._transfer_callback = cb

    def set_source_text(self, text: str):
        self.src_edit.setPlainText(text)

    def _apply(self):
        text = self.src_edit.toPlainText().strip()
        if not text:
            self.status_bar.showMessage('Введите исходный текст')
            return
        mode = self.mode_combo.currentIndex()
        try:
            if mode == 0:
                result = self.modifier.replace_synonyms(text)
            elif mode == 1:
                api_key = self.api_key_input.text().strip()
                result = self.modifier.paraphrase_gigachat(text, api_key)
            else:
                result = self.modifier.diversify_sentences(text)
            self.dst_edit.setPlainText(result)
            self.status_bar.showMessage('Модификация выполнена')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

    def _copy_result(self):
        from PyQt5.QtWidgets import QApplication
        text = self.dst_edit.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            self.status_bar.showMessage('Скопировано в буфер обмена')

    def _request_transfer(self):
        if self._transfer_callback:
            text = self._transfer_callback()
            if text:
                self.src_edit.setPlainText(text)
                self.status_bar.showMessage('Текст перенесён из вкладки Анализ')
            else:
                self.status_bar.showMessage('Текст на вкладке Анализ пуст')


# ════════════════════════════════════════════════════════════════════════════
#  Tab 4 – Phrase DB
# ════════════════════════════════════════════════════════════════════════════
class PhrasesTab(QWidget):
    def __init__(self, phrase_db: PhraseDB, status_bar: QStatusBar):
        super().__init__()
        self.phrase_db = phrase_db
        self.status_bar = status_bar
        self._build_ui()

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # List
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.setSpacing(6)

        lv.addWidget(QLabel('Шаблонные фразы:'))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Поиск...')
        self.search_input.textChanged.connect(self._search)
        lv.addWidget(self.search_input)

        self.list_widget = QListWidget()
        lv.addWidget(self.list_widget)
        root.addWidget(left, stretch=3)

        # Buttons
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(8)
        rv.setAlignment(Qt.AlignTop)

        for label, slot in (
            ('Добавить',             self._add),
            ('Редактировать',        self._edit),
            ('Удалить',              self._delete),
            ('Сбросить к умолчаниям', self._reset),
        ):
            btn = QPushButton(label)
            btn.setFixedHeight(32)
            btn.clicked.connect(slot)
            rv.addWidget(btn)

        self.count_label = QLabel()
        self.count_label.setStyleSheet(f'color: {TEXT_DIM}; font-size: 11px;')
        rv.addWidget(self.count_label)

        root.addWidget(right, stretch=1)
        self._refresh()

    def _refresh(self, phrases=None):
        if phrases is None:
            phrases = self.phrase_db.get_all()
        self.list_widget.clear()
        for p in phrases:
            self.list_widget.addItem(p)
        self.count_label.setText(f'Фраз: {self.phrase_db.count()}')

    def _search(self, query: str):
        if query.strip():
            self._refresh(self.phrase_db.search(query))
        else:
            self._refresh()

    def _add(self):
        text, ok = QInputDialog.getText(self, 'Добавить фразу', 'Новая фраза:')
        if ok and text.strip():
            self.phrase_db.add(text.strip())
            self._refresh()
            self.status_bar.showMessage('Фраза добавлена')

    def _edit(self):
        item = self.list_widget.currentItem()
        if not item:
            self.status_bar.showMessage('Выберите фразу для редактирования')
            return
        old = item.text()
        text, ok = QInputDialog.getText(self, 'Редактировать фразу', 'Фраза:', text=old)
        if ok and text.strip():
            self.phrase_db.update(old, text.strip())
            self._refresh()
            self.status_bar.showMessage('Фраза обновлена')

    def _delete(self):
        item = self.list_widget.currentItem()
        if not item:
            self.status_bar.showMessage('Выберите фразу для удаления')
            return
        reply = QMessageBox.question(
            self, 'Удалить', f'Удалить фразу?\n"{item.text()}"',
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.phrase_db.remove(item.text())
            self._refresh()
            self.status_bar.showMessage('Фраза удалена')

    def _reset(self):
        reply = QMessageBox.question(
            self, 'Сброс', 'Сбросить базу к умолчаниям?',
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.phrase_db.reset_to_defaults()
            self._refresh()
            self.status_bar.showMessage('База фраз сброшена к умолчаниям')


# ════════════════════════════════════════════════════════════════════════════
#  Main Window
# ════════════════════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('TextAnalyzer Pro')
        self.resize(1200, 750)
        self.setMinimumSize(900, 600)
        self.setStyleSheet(STYLESHEET)

        self.phrase_db = PhraseDB(
            os.path.join(os.path.dirname(os.path.dirname(__file__)), 'phrases.json')
        )

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage('TextAnalyzer Pro готов к работе')

        tabs = QTabWidget()
        tabs.setDocumentMode(True)

        self.analysis_tab = AnalysisTab(self.phrase_db, self.status_bar)
        self.compare_tab  = CompareTab(self.status_bar)
        self.modify_tab   = ModifyTab(self.status_bar)
        self.phrases_tab  = PhrasesTab(self.phrase_db, self.status_bar)

        # Wire up transfer from Analysis -> Modify
        self.modify_tab.set_transfer_callback(self.analysis_tab.get_text)

        tabs.addTab(self.analysis_tab, 'Анализ')
        tabs.addTab(self.compare_tab,  'Сравнение')
        tabs.addTab(self.modify_tab,   'Модификация')
        tabs.addTab(self.phrases_tab,  'База фраз')

        self.setCentralWidget(tabs)
