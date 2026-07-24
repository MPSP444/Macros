import sys
import subprocess
import platform
import os
import re
import time
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
}


def run_powershell(cmd: str) -> str:
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True, text=True, timeout=30
        )
        return result.stdout.strip()
    except Exception as e:
        return f"ERROR: {e}"


class CheckWorker(QThread):
    progress = pyqtSignal(str, int)
    row_ready = pyqtSignal(str, str, str, str)  # категория, ожидается, обнаружено, статус
    finished_all = pyqtSignal()

    def run(self):
        self.progress.emit("Проверка процессора...", 10)
        self.check_cpu()
        self.progress.emit("Проверка оперативной памяти...", 25)
        self.check_ram()
        self.progress.emit("Проверка видеокарты...", 40)
        self.check_gpu()
        self.progress.emit("Проверка накопителя...", 55)
        self.check_disk()
        self.progress.emit("Тест скорости SSD (может занять ~30 сек)...", 70)
        self.check_disk_speed()
        self.progress.emit("Проверка экрана...", 85)
        self.check_display()
        self.progress.emit("Проверка батареи...", 95)
        self.check_battery()
        self.progress.emit("Готово", 100)
        self.finished_all.emit()

    def check_cpu(self):
        actual = run_powershell("(Get-CimInstance Win32_Processor).Name")
        expected = EXPECTED_SPECS["cpu_model"]
        status = "OK" if expected.lower() in actual.lower() else "НЕСОВПАДЕНИЕ"
        self.row_ready.emit("Процессор (модель)", expected, actual or "не определено", status)

        try:
            cores = int(run_powershell("(Get-CimInstance Win32_Processor).NumberOfCores"))
        except ValueError:
            cores = None
        try:
            threads = int(run_powershell("(Get-CimInstance Win32_Processor).NumberOfLogicalProcessors"))
        except ValueError:
            threads = None

        self.row_ready.emit("Ядра CPU", str(EXPECTED_SPECS["cpu_cores"]), str(cores),
                             "OK" if cores == EXPECTED_SPECS["cpu_cores"] else "ПРОВЕРЬ")
        self.row_ready.emit("Потоки CPU", str(EXPECTED_SPECS["cpu_threads"]), str(threads),
                             "OK" if threads == EXPECTED_SPECS["cpu_threads"] else "ПРОВЕРЬ")

    def check_ram(self):
        out = run_powershell("[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB,1)")
        try:
            actual_gb = float(out)
        except ValueError:
            actual_gb = None
        expected = EXPECTED_SPECS["ram_gb"]
        status = "OK" if actual_gb and abs(actual_gb - expected) <= 1 else "НЕСОВПАДЕНИЕ"
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

    def check_disk_speed(self):
        test_file = os.path.join(tempfile.gettempdir(), "disk_speed_test.tmp")
        size_mb = 500
        chunk = os.urandom(1024 * 1024)
        try:
            start = time.time()
            with open(test_file, "wb") as f:
                for _ in range(size_mb):
                    f.write(chunk)
                f.flush()
                os.fsync(f.fileno())
            write_speed = size_mb / (time.time() - start)

            start = time.time()
            with open(test_file, "rb") as f:
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
        expected = f"{EXPECTED_SPECS['screen_resolution'][0]} x {EXPECTED_SPECS['screen_resolution'][1]}"
        self.row_ready.emit("Разрешение экрана", expected, out or "не определено", "ИНФО")

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
        self.resize(900, 600)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        title = QLabel("Проверка соответствия характеристик ноутбука")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title)

        subtitle = QLabel("Ожидается: AMD Ryzen AI 7 450 · Radeon 860M · 32 ГБ RAM · 1 ТБ SSD · 16\" 3200x2000")
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
        layout.addWidget(self.summary_label)

        self.worker = None
        self.mismatch_count = 0
        self.total_checked = 0

    def start_check(self):
        self.table.setRowCount(0)
        self.mismatch_count = 0
        self.total_checked = 0
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

        if status == "OK":
            status_item.setBackground(QColor(200, 255, 200))
        elif status in ("НЕСОВПАДЕНИЕ", "МЕДЛЕННО", "ОШИБКА"):
            status_item.setBackground(QColor(255, 200, 200))
            self.mismatch_count += 1
        elif status == "ПРОВЕРЬ":
            status_item.setBackground(QColor(255, 235, 180))
        self.table.setItem(row, 3, status_item)

        if status != "ИНФО":
            self.total_checked += 1

    def on_finished(self):
        self.status_label.setText("Проверка завершена.")
        self.run_btn.setEnabled(True)
        self.export_btn.setEnabled(True)

        if self.mismatch_count == 0:
            self.summary_label.setText("✅ Явных несоответствий не найдено.")
            self.summary_label.setStyleSheet("color: green;")
        else:
            self.summary_label.setText(
                f"⚠️ Найдено {self.mismatch_count} несоответствий из {self.total_checked} — смотри красные строки."
            )
            self.summary_label.setStyleSheet("color: red;")

    def export_report(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", "laptop_check_report.txt", "Текстовые файлы (*.txt)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("Отчёт о проверке ноутбука\n" + "=" * 50 + "\n\n")
            for row in range(self.table.rowCount()):
                cat = self.table.item(row, 0).text()
                exp = self.table.item(row, 1).text()
                act = self.table.item(row, 2).text()
                stat = self.table.item(row, 3).text()
                f.write(f"{cat}: ожидается [{exp}] / обнаружено [{act}] — {stat}\n")
        QMessageBox.information(self, "Готово", f"Отчёт сохранён: {path}")


def main():
    if platform.system() != "Windows":
        print("Внимание: полная проверка (батарея, WMI) работает только на Windows.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()