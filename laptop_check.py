import sys
import subprocess
import platform
import os
import re
import time
import shutil
import tempfile

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QProgressBar,
    QHeaderView, QMessageBox, QFileDialog
)
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont

# Ожидаемая конфигурация — поменяй под свою модель, если понадобится
EXPECTED_SPECS = {
    "cpu_model": "Ryzen AI 7 450",
    "cpu_cores": 8,
    "cpu_threads": 16,
    "gpu_model": "860M",
    "ram_gb": 32,
    "ssd_gb": 1024,
    "screen_resolution": (3200, 2000),
    "disk_bus_type": "NVMe",
    "refresh_rate_hz": 165,
    "wifi_keywords": ["AX211", "MT7925", "BE200", "BE201", "BE202"],
}

# Категории, которые считаются ключевыми при формировании итоговой рекомендации
CRITICAL_CATEGORIES = {"Процессор (модель)", "Объём RAM (ГБ)", "Видеокарта"}

# Единая палитра статусов — используется и в таблице интерфейса, и в HTML-экспорте
STATUS_COLORS = {
    "OK": "#c8ffc8",
    "НЕСОВПАДЕНИЕ": "#ffc8c8",
    "МЕДЛЕННО": "#ffc8c8",
    "ОШИБКА": "#ffc8c8",
    "ПРОВЕРЬ": "#ffebb4",
    "ИНФО": "#f0f0f0",
}


def run_powershell(cmd: str) -> str:
    # Числа/даты в PowerShell форматируются по локали ОС — на русской Windows
    # разделитель дробной части запятая, что ломает float()/int() на стороне Python.
    # Принудительно переключаем поток PowerShell на InvariantCulture перед выполнением.
    culture_prefix = (
        "[System.Threading.Thread]::CurrentThread.CurrentCulture = "
        "[System.Globalization.CultureInfo]::InvariantCulture; "
        "[System.Threading.Thread]::CurrentThread.CurrentUICulture = "
        "[System.Globalization.CultureInfo]::InvariantCulture; "
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", culture_prefix + cmd],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout.strip()
    except Exception as e:
        return f"ERROR: {e}"


def parse_float(s):
    # Доп. защита на случай, если культура PowerShell всё же не применилась
    # (например, старая версия PS) — заменяем запятую на точку перед float().
    if not s:
        return None
    try:
        return float(s.strip().replace(",", "."))
    except ValueError:
        return None


def parse_int(s):
    value = parse_float(s)
    return int(value) if value is not None else None


class CheckWorker(QThread):
    progress = pyqtSignal(str, int)
    row_ready = pyqtSignal(str, str, str, str)  # категория, ожидается, обнаружено, статус
    finished_all = pyqtSignal()

    def run(self):
        self.progress.emit("Проверка процессора...", 5)
        self.check_cpu()
        self.progress.emit("Проверка оперативной памяти...", 15)
        self.check_ram()
        self.progress.emit("Проверка видеокарты...", 22)
        self.check_gpu()
        self.progress.emit("Проверка накопителя...", 30)
        self.check_disk()
        self.progress.emit("Проверка типа накопителя (NVMe/SATA)...", 38)
        self.check_disk_type()
        self.progress.emit("Тест скорости накопителя (может занять 1-2 мин)...", 55)
        self.check_disk_speed()
        self.progress.emit("Проверка экрана...", 65)
        self.check_display()
        self.progress.emit("Проверка Wi-Fi модуля...", 72)
        self.check_wifi()
        self.progress.emit("Проверка серийного номера и даты производства...", 78)
        self.check_serial()
        self.progress.emit("Проверка активации Windows...", 84)
        self.check_windows_activation()
        self.progress.emit("Проверка NPU...", 89)
        self.check_npu()
        self.progress.emit("Проверка батареи...", 95)
        self.check_battery()
        self.progress.emit("Готово", 100)
        self.finished_all.emit()

    def check_cpu(self):
        actual = run_powershell("(Get-CimInstance Win32_Processor).Name")
        expected = EXPECTED_SPECS["cpu_model"]
        status = "OK" if expected.lower() in actual.lower() else "НЕСОВПАДЕНИЕ"
        self.row_ready.emit("Процессор (модель)", expected, actual or "не определено", status)

        cores = parse_int(run_powershell("(Get-CimInstance Win32_Processor).NumberOfCores"))
        threads = parse_int(run_powershell("(Get-CimInstance Win32_Processor).NumberOfLogicalProcessors"))

        self.row_ready.emit("Ядра CPU", str(EXPECTED_SPECS["cpu_cores"]), str(cores),
                             "OK" if cores == EXPECTED_SPECS["cpu_cores"] else "ПРОВЕРЬ")
        self.row_ready.emit("Потоки CPU", str(EXPECTED_SPECS["cpu_threads"]), str(threads),
                             "OK" if threads == EXPECTED_SPECS["cpu_threads"] else "ПРОВЕРЬ")

    def check_ram(self):
        out = run_powershell("[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB,1)")
        actual_gb = parse_float(out)
        expected = EXPECTED_SPECS["ram_gb"]
        status = "OK" if actual_gb is not None and abs(actual_gb - expected) <= 1 else "НЕСОВПАДЕНИЕ"
        self.row_ready.emit("Объём RAM (ГБ)", str(expected), str(actual_gb), status)

        speed = run_powershell("(Get-CimInstance Win32_PhysicalMemory | Select -First 1).Speed")
        self.row_ready.emit("Частота RAM (МГц)", "сверь на сайте Lenovo", speed or "не определено", "ИНФО")

    def check_gpu(self):
        actual = run_powershell("(Get-CimInstance Win32_VideoController).Name")
        status = "OK" if EXPECTED_SPECS["gpu_model"] in actual else "ПРОВЕРЬ"
        self.row_ready.emit("Видеокарта", f"...{EXPECTED_SPECS['gpu_model']}", actual or "не определено", status)

    def check_disk(self):
        out = run_powershell(
            "Get-CimInstance Win32_DiskDrive | "
            "Select-Object Model, @{N='SizeGB';E={[math]::Round($_.Size/1GB,0)}} | "
            "Format-Table -HideTableHeaders"
        )
        self.row_ready.emit("Диск (модель/объём)", f"~{EXPECTED_SPECS['ssd_gb']} ГБ", out or "не определено", "ИНФО")

    def check_disk_type(self):
        out = run_powershell(
            "Get-PhysicalDisk | Select-Object FriendlyName, BusType | "
            "Format-Table -HideTableHeaders | Out-String -Width 300"
        )
        expected = EXPECTED_SPECS["disk_bus_type"]
        text = (out or "").strip()
        lowered = text.lower()
        if "sata" in lowered:
            status = "НЕСОВПАДЕНИЕ"
        elif expected.lower() in lowered:
            status = "OK"
        else:
            status = "ПРОВЕРЬ"
        self.row_ready.emit("Тип накопителя (шина)", expected, text or "не определено", status)

    def check_disk_speed(self):
        # Опционально: если доступна портативная CrystalDiskMark CLI (путь можно
        # передать через переменную окружения CDM_CLI_PATH), используем её для
        # честного теста. Формат вывода отличается между версиями CDM, поэтому
        # парсинг best-effort — при неудаче автоматически откатываемся на файловый тест.
        cdm_path = os.environ.get("CDM_CLI_PATH") or shutil.which("DiskMarkCmd") or shutil.which("DiskMark64")
        if cdm_path and os.path.exists(cdm_path) and self._run_crystaldiskmark(cdm_path):
            return
        self._check_disk_speed_fallback()

    def _run_crystaldiskmark(self, cdm_path):
        try:
            extra_args = os.environ.get("CDM_CLI_ARGS", "").split()
            result = subprocess.run(
                [cdm_path] + extra_args, capture_output=True, text=True, timeout=180
            )
            output = (result.stdout or "") + (result.stderr or "")
            read_match = re.search(r"seq.*?read.*?([\d.,]+)\s*MB/s", output, re.I)
            write_match = re.search(r"seq.*?write.*?([\d.,]+)\s*MB/s", output, re.I)
            if not (read_match and write_match):
                return False
            read_speed = parse_float(read_match.group(1))
            write_speed = parse_float(write_match.group(1))
            if read_speed is None or write_speed is None:
                return False
            self.row_ready.emit("Скорость чтения CrystalDiskMark (МБ/с)", "> 2000 (NVMe Gen4)",
                                 f"{read_speed:.0f}", "OK" if read_speed > 300 else "МЕДЛЕННО")
            self.row_ready.emit("Скорость записи CrystalDiskMark (МБ/с)", "> 1000 (NVMe Gen4)",
                                 f"{write_speed:.0f}", "OK" if write_speed > 300 else "МЕДЛЕННО")
            return True
        except Exception:
            return False

    def _check_disk_speed_fallback(self):
        test_file = os.path.join(tempfile.gettempdir(), "disk_speed_test.tmp")
        size_mb = 2048  # 2 ГБ — крупнее файл даёт более стабильный результат, кэш меньше искажает цифры
        chunk = os.urandom(1024 * 1024)
        try:
            start = time.time()
            with open(test_file, "wb", buffering=0) as f:
                for _ in range(size_mb):
                    f.write(chunk)
                os.fsync(f.fileno())
            write_speed = size_mb / (time.time() - start)

            start = time.time()
            with open(test_file, "rb", buffering=0) as f:
                while f.read(1024 * 1024):
                    pass
            read_speed = size_mb / (time.time() - start)
            os.remove(test_file)

            self.row_ready.emit("Скорость записи (МБ/с)", "> 1000 (NVMe Gen4)", f"{write_speed:.0f}",
                                 "OK" if write_speed > 300 else "МЕДЛЕННО")
            self.row_ready.emit("Скорость чтения (МБ/с)", "> 2000 (NVMe Gen4)", f"{read_speed:.0f}",
                                 "OK" if read_speed > 300 else "МЕДЛЕННО")
        except Exception as e:
            self.row_ready.emit("Тест скорости диска", "-", f"ошибка: {e}", "ОШИБКА")

    def check_display(self):
        out = run_powershell(
            "Get-CimInstance Win32_VideoController | "
            "Select-Object CurrentHorizontalResolution, CurrentVerticalResolution | "
            "Format-Table -HideTableHeaders"
        )
        expected_res = f"{EXPECTED_SPECS['screen_resolution'][0]} x {EXPECTED_SPECS['screen_resolution'][1]}"
        self.row_ready.emit("Разрешение экрана", expected_res, out or "не определено", "ИНФО")

        refresh_out = run_powershell("(Get-CimInstance Win32_VideoController | Select-Object -First 1).MaxRefreshRate")
        expected_hz = EXPECTED_SPECS["refresh_rate_hz"]
        refresh = parse_int(refresh_out)
        if refresh is None:
            status = "ИНФО"
            actual_text = refresh_out or "не определено"
        else:
            status = "OK" if refresh >= expected_hz else "ПРОВЕРЬ"
            actual_text = f"{refresh} Гц (текущий режим Windows — сверь также в Параметры → Экран → Доп. параметры)"
        self.row_ready.emit("Частота обновления экрана", f"{expected_hz} Гц", actual_text, status)

    def check_wifi(self):
        out = run_powershell(
            "(Get-NetAdapter | Where-Object {$_.Name -like '*Wi-Fi*' -or $_.InterfaceDescription -like '*Wireless*'} | "
            "Select-Object -First 1).InterfaceDescription"
        )
        keywords = EXPECTED_SPECS["wifi_keywords"]
        matched = any(k.lower() in (out or "").lower() for k in keywords)
        status = "OK" if matched else "ПРОВЕРЬ"
        self.row_ready.emit("Wi-Fi модуль", f"Wi-Fi 7 (например: {', '.join(keywords)})",
                             out or "не определено", status)

    def check_serial(self):
        serial = run_powershell("(Get-CimInstance Win32_BIOS).SerialNumber")
        self.row_ready.emit(
            "Серийный номер",
            "введи на support.lenovo.com для проверки комплектации",
            serial or "не определено",
            "ИНФО",
        )
        release = run_powershell("(Get-CimInstance Win32_BIOS).ReleaseDate.ToString('yyyy-MM-dd')")
        self.row_ready.emit("Дата производства (BIOS)", "сверь дату выпуска модели", release or "не определено", "ИНФО")

    def check_windows_activation(self):
        out = run_powershell(
            "(Get-CimInstance SoftwareLicensingProduct -Filter "
            "\"ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL\").LicenseStatus"
        )
        status_map = {
            "0": "не активирована",
            "1": "лицензирована",
            "2": "льготный период (OOB Grace)",
            "3": "льготный период (OOT Grace)",
            "4": "не оригинальная (Non-Genuine Grace)",
            "5": "уведомление",
            "6": "продлённый льготный период",
        }
        code = (out or "").strip().splitlines()[0].strip() if out else ""
        human = status_map.get(code, f"неизвестно ({out})" if out else "не удалось определить")
        status = "OK" if code == "1" else ("НЕСОВПАДЕНИЕ" if code else "ПРОВЕРЬ")
        self.row_ready.emit("Активация Windows", "лицензирована", human, status)

    def check_npu(self):
        out = run_powershell(
            "(Get-PnpDevice | Where-Object {$_.Class -eq 'SoftwareComponent' -or $_.FriendlyName -like '*NPU*' "
            "-or $_.FriendlyName -like '*Neural*'} | Select-Object -ExpandProperty FriendlyName) -join '; '"
        )
        status = "OK" if out else "ПРОВЕРЬ"
        self.row_ready.emit("NPU (нейропроцессор)", "присутствует (Ryzen AI, до 50 TOPS)", out or "не обнаружено", status)

    def check_battery(self):
        report_path = os.path.join(tempfile.gettempdir(), "battery_report.html")
        run_powershell(f'powercfg /batteryreport /output "{report_path}"')
        design_cap = full_cap = cycle_count = None
        if os.path.exists(report_path):
            try:
                html = open(report_path, "r", encoding="utf-8", errors="ignore").read()
                m1 = re.search(r"DESIGN CAPACITY</td>\s*<td[^>]*>([\d,]+)\s*mWh", html, re.I)
                m2 = re.search(r"FULL CHARGE CAPACITY</td>\s*<td[^>]*>([\d,]+)\s*mWh", html, re.I)
                m3 = re.search(r"CYCLE COUNT</td>\s*<td[^>]*>(\d+)", html, re.I)
                if m1: design_cap = int(m1.group(1).replace(",", ""))
                if m2: full_cap = int(m2.group(1).replace(",", ""))
                if m3: cycle_count = int(m3.group(1))
            except Exception:
                pass

        if design_cap and full_cap:
            health = round(full_cap / design_cap * 100, 1)
            self.row_ready.emit("Здоровье батареи (%)", "≥ 95% (новый)", f"{health}%",
                                 "OK" if health >= 95 else "ПРОВЕРЬ")
        else:
            self.row_ready.emit("Здоровье батареи", "≥ 95% (новый)", "не удалось прочитать отчёт", "ИНФО")

        if cycle_count is not None:
            self.row_ready.emit("Циклы зарядки", "0-5 (новый)", str(cycle_count),
                                 "OK" if cycle_count <= 5 else "ПРОВЕРЬ")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Проверка ноутбука — Lenovo ThinkBook 16+ 2026")
        self.resize(950, 650)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        title = QLabel("Проверка соответствия характеристик ноутбука")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title)

        subtitle = QLabel(
            "Ожидается: AMD Ryzen AI 7 450 · Radeon 860M · 32 ГБ RAM · 1 ТБ SSD (NVMe) · "
            "16\" 3200x2000 165Гц · Wi-Fi 7"
        )
        subtitle.setStyleSheet("color: gray;")
        layout.addWidget(subtitle)

        btn_layout = QHBoxLayout()
        self.run_btn = QPushButton("Запустить полную проверку")
        self.run_btn.clicked.connect(self.start_check)
        btn_layout.addWidget(self.run_btn)

        self.export_btn = QPushButton("Сохранить отчёт в файл")
        self.export_btn.clicked.connect(self.export_report)
        self.export_btn.setEnabled(False)
        btn_layout.addWidget(self.export_btn)
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Нажми «Запустить полную проверку», чтобы начать.")
        layout.addWidget(self.status_label)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Параметр", "Ожидается", "Обнаружено", "Статус"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.summary_label = QLabel("")
        self.summary_label.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.worker = None
        self.mismatch_count = 0
        self.total_checked = 0
        self.field_status = {}

    def start_check(self):
        self.table.setRowCount(0)
        self.mismatch_count = 0
        self.total_checked = 0
        self.field_status = {}
        self.summary_label.setText("")
        self.run_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress_bar.setValue(0)

        self.worker = CheckWorker()
        self.worker.progress.connect(self.on_progress)
        self.worker.row_ready.connect(self.on_row_ready)
        self.worker.finished_all.connect(self.on_finished)
        self.worker.start()

    def on_progress(self, message, percent):
        self.status_label.setText(message)
        self.progress_bar.setValue(percent)

    def on_row_ready(self, category, expected, actual, status):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(category))
        self.table.setItem(row, 1, QTableWidgetItem(expected))
        self.table.setItem(row, 2, QTableWidgetItem(actual))
        status_item = QTableWidgetItem(status)

        color_hex = STATUS_COLORS.get(status)
        if color_hex:
            status_item.setBackground(QColor(color_hex))
        self.table.setItem(row, 3, status_item)

        self.field_status[category] = status

        if status in ("НЕСОВПАДЕНИЕ", "МЕДЛЕННО", "ОШИБКА"):
            self.mismatch_count += 1

        if status != "ИНФО":
            self.total_checked += 1

    def compute_conclusion(self):
        critical_mismatch = any(
            self.field_status.get(cat) == "НЕСОВПАДЕНИЕ" for cat in CRITICAL_CATEGORIES
        )
        if self.mismatch_count == 0:
            return "Конфигурация соответствует заявленной. Проблем не найдено.", "green"
        elif self.mismatch_count <= 2 and not critical_mismatch:
            return ("Мелкие расхождения, но ключевые компоненты (CPU/RAM/GPU) совпадают — "
                    "вероятно, всё в порядке."), "#b8860b"
        else:
            return ("⚠️ КРИТИЧНО: ключевые компоненты не соответствуют заявленным. "
                    "Рекомендуется отказаться от покупки/оформить возврат."), "red"

    def on_finished(self):
        self.status_label.setText("Проверка завершена.")
        self.run_btn.setEnabled(True)
        self.export_btn.setEnabled(True)

        conclusion, color = self.compute_conclusion()
        if self.mismatch_count == 0:
            header = "✅ Явных несоответствий не найдено."
        else:
            header = f"⚠️ Найдено {self.mismatch_count} несоответствий из {self.total_checked} — смотри красные строки."
        self.summary_label.setText(f"{header}\n{conclusion}")
        self.summary_label.setStyleSheet(f"color: {color};")

    def export_report(self):
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", "laptop_check_report.html",
            "HTML отчёт (*.html);;Текстовый отчёт (*.txt)"
        )
        if not path:
            return

        is_html = path.lower().endswith(".html") or "html" in selected_filter.lower()
        if is_html and not path.lower().endswith(".html"):
            path += ".html"
        elif not is_html and not path.lower().endswith(".txt"):
            path += ".txt"

        conclusion, _ = self.compute_conclusion()
        if is_html:
            self._export_html(path, conclusion)
        else:
            self._export_txt(path, conclusion)

        QMessageBox.information(self, "Готово", f"Отчёт сохранён: {path}")

    def _export_txt(self, path, conclusion):
        with open(path, "w", encoding="utf-8") as f:
            f.write("Отчёт о проверке ноутбука\n" + "=" * 50 + "\n\n")
            for row in range(self.table.rowCount()):
                cat = self.table.item(row, 0).text()
                exp = self.table.item(row, 1).text()
                act = self.table.item(row, 2).text()
                stat = self.table.item(row, 3).text()
                f.write(f"{cat}: ожидается [{exp}] / обнаружено [{act}] — {stat}\n")
            f.write("\n" + "=" * 50 + "\n")
            f.write(f"Итог: {conclusion}\n")

    def _export_html(self, path, conclusion):
        def esc(s):
            return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        rows_html = []
        for row in range(self.table.rowCount()):
            cat = self.table.item(row, 0).text()
            exp = self.table.item(row, 1).text()
            act = self.table.item(row, 2).text()
            stat = self.table.item(row, 3).text()
            color = STATUS_COLORS.get(stat, "#ffffff")
            rows_html.append(
                f"<tr style='background:{color}'>"
                f"<td>{esc(cat)}</td><td>{esc(exp)}</td><td>{esc(act)}</td><td>{esc(stat)}</td></tr>"
            )

        html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<title>Отчёт о проверке ноутбука</title>
<style>
  body {{ font-family: "Segoe UI", Arial, sans-serif; margin: 24px; color: #222; }}
  table {{ border-collapse: collapse; width: 100%; }}
  th, td {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left; }}
  th {{ background: #eee; }}
  .conclusion {{ margin-top: 20px; padding: 12px; font-weight: bold; border-radius: 4px; background: #f4f4f4; }}
</style>
</head>
<body>
<h2>Отчёт о проверке ноутбука</h2>
<table>
<tr><th>Параметр</th><th>Ожидается</th><th>Обнаружено</th><th>Статус</th></tr>
{''.join(rows_html)}
</table>
<div class="conclusion">Итог: {esc(conclusion)}</div>
</body>
</html>
"""
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)


def main():
    if platform.system() != "Windows":
        print("Внимание: полная проверка (батарея, WMI) работает только на Windows.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
