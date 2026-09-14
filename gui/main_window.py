"""
MainWindow - satu-satunya tempat GUI "berbicara" dengan backend.
GUI hanya memanggil ReportProcessor lewat ProcessWorker; tidak ada
logika bisnis (mapping, dedup, dsb.) di file ini.
"""

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QWidget,
    QMainWindow,
    QLabel,
    QLineEdit,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QFileDialog,
    QMessageBox,
    QGroupBox,
)

from gui.worker import ProcessWorker


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Laporan Harian Otomatis")
        self.resize(820, 640)

        self.worker = None
        self.last_output_folder = None

        self._build_ui()

    # =====================================
    # BUILD UI
    # =====================================

    def _build_ui(self):

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # ---------- Input Group ----------
        input_box = QGroupBox("Input")
        grid = QGridLayout(input_box)

        self.zip_folder_edit = QLineEdit()
        self.zip_folder_edit.setPlaceholderText("Folder yang berisi file-file ZIP...")
        zip_browse_btn = QPushButton("Browse Folder ZIP")
        zip_browse_btn.clicked.connect(self._browse_zip_folder)

        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("File Template Excel (Master)...")
        template_browse_btn = QPushButton("Browse Template Excel")
        template_browse_btn.clicked.connect(self._browse_template)

        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("Folder tujuan output...")
        output_browse_btn = QPushButton("Browse Folder Output")
        output_browse_btn.clicked.connect(self._browse_output_folder)

        grid.addWidget(QLabel("Folder ZIP"), 0, 0)
        grid.addWidget(self.zip_folder_edit, 0, 1)
        grid.addWidget(zip_browse_btn, 0, 2)

        grid.addWidget(QLabel("Template Excel"), 1, 0)
        grid.addWidget(self.template_edit, 1, 1)
        grid.addWidget(template_browse_btn, 1, 2)

        grid.addWidget(QLabel("Folder Output"), 2, 0)
        grid.addWidget(self.output_folder_edit, 2, 1)
        grid.addWidget(output_browse_btn, 2, 2)

        root.addWidget(input_box)

        # ---------- Action Row ----------
        action_row = QHBoxLayout()

        self.process_btn = QPushButton("Process")
        self.process_btn.setMinimumHeight(36)
        self.process_btn.clicked.connect(self._start_process)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setMinimumHeight(36)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_process)

        self.open_output_btn = QPushButton("Open Output Folder")
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.clicked.connect(self._open_output_folder)

        action_row.addWidget(self.process_btn)
        action_row.addWidget(self.stop_btn)
        action_row.addWidget(self.open_output_btn)
        root.addLayout(action_row)

        # ---------- Progress ----------
        self.status_label = QLabel("Siap.")
        root.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)

        # ---------- Log ----------
        log_box = QGroupBox("Log Proses")
        log_layout = QVBoxLayout(log_box)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_view)
        root.addWidget(log_box, stretch=1)

        # ---------- Summary ----------
        summary_box = QGroupBox("Ringkasan Hasil")
        summary_layout = QVBoxLayout(summary_box)
        self.summary_label = QLabel("Belum ada proses yang dijalankan.")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)
        root.addWidget(summary_box)

    # =====================================
    # BROWSE HANDLERS
    # =====================================

    def _browse_zip_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Pilih Folder ZIP")
        if folder:
            self.zip_folder_edit.setText(folder)

    def _browse_template(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Pilih Template Excel", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            self.template_edit.setText(file_path)

    def _browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
        if folder:
            self.output_folder_edit.setText(folder)

    def _open_output_folder(self):
        if self.last_output_folder:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_output_folder)))

    # =====================================
    # VALIDATION + START
    # =====================================

    def _start_process(self):

        zip_folder = self.zip_folder_edit.text().strip()
        template_file = self.template_edit.text().strip()
        output_folder = self.output_folder_edit.text().strip()

        if not zip_folder or not Path(zip_folder).is_dir():
            QMessageBox.warning(self, "Input Tidak Lengkap", "Folder ZIP belum dipilih atau tidak valid.")
            return

        if not template_file or not Path(template_file).is_file():
            QMessageBox.warning(self, "Input Tidak Lengkap", "Template Excel belum dipilih atau tidak valid.")
            return

        if not output_folder:
            QMessageBox.warning(self, "Input Tidak Lengkap", "Folder Output belum dipilih.")
            return

        # rglob (bukan glob) supaya ZIP di dalam subfolder (mis. APRIL/PT ABC/xxx.zip)
        # ikut ditemukan, tidak cuma yang persis di folder teratas.
        zip_paths = sorted(Path(zip_folder).rglob("*.zip"))

        if not zip_paths:
            QMessageBox.warning(self, "Tidak Ada ZIP", "Tidak ditemukan file .zip di folder tersebut.")
            return

        self._set_running_state(True)
        self.log_view.clear()
        self.summary_label.setText("Memproses...")
        self.progress_bar.setValue(0)
        self.status_label.setText(f"Memproses {len(zip_paths)} file ZIP...")

        self.log_view.appendPlainText(f"Ditemukan {len(zip_paths)} file ZIP:")
        for zp in zip_paths:
            self.log_view.appendPlainText(f"  - {zp.relative_to(zip_folder)}")
        self.log_view.appendPlainText("")

        self.worker = ProcessWorker(
            zip_paths=zip_paths,
            master_file=template_file,
            output_folder=output_folder,
        )

        self.worker.log_line.connect(self._on_log)
        self.worker.zip_progress.connect(self._on_zip_progress)
        self.worker.file_progress.connect(self._on_file_progress)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.stopped.connect(self._on_stopped)
        self.worker.failed.connect(self._on_failed)

        self._zip_total = len(zip_paths)
        self.worker.start()

    def _stop_process(self):
        if self.worker:
            self.stop_btn.setEnabled(False)
            self.status_label.setText("Menghentikan (menunggu file yang sedang berjalan selesai)...")
            self.worker.request_stop()

    def _set_running_state(self, running):
        self.process_btn.setEnabled(not running)
        self.stop_btn.setEnabled(running)
        self.open_output_btn.setEnabled(False if running else self.open_output_btn.isEnabled())

    # =====================================
    # WORKER CALLBACKS
    # =====================================

    def _on_log(self, message):
        self.log_view.appendPlainText(message)

    def _on_zip_progress(self, index, total, nama_zip):
        self.status_label.setText(f"ZIP {index}/{total}: {nama_zip}")
        # progress kasar di level ZIP; progress detail per-file dari _on_file_progress
        overall = int((index - 1) / max(total, 1) * 100)
        self.progress_bar.setValue(overall)

    def _on_file_progress(self, index, total):
        if total:
            self.progress_bar.setValue(int(index / total * 100))

    def _on_finished_ok(self, output_path):
        self._set_running_state(False)
        self.progress_bar.setValue(100)

        output_path = Path(output_path)
        self.last_output_folder = output_path.parent
        self.open_output_btn.setEnabled(True)

        self.status_label.setText("Selesai.")
        self.summary_label.setText(
            f"Proses selesai.\nFile output: {output_path.name}\nLokasi: {output_path.parent}"
        )

    def _on_stopped(self, output_path):
        self._set_running_state(False)

        output_path = Path(output_path)
        self.last_output_folder = output_path.parent
        self.open_output_btn.setEnabled(True)

        self.status_label.setText("Dihentikan oleh pengguna.")
        self.summary_label.setText(
            f"Proses dihentikan sebelum selesai.\n"
            f"Data yang sempat digabung tetap disimpan di: {output_path.name}\n"
            f"Lokasi: {output_path.parent}"
        )

    def _on_failed(self, error_message):
        self._set_running_state(False)
        self.status_label.setText("Gagal.")
        self.summary_label.setText(f"Proses gagal: {error_message}")
        QMessageBox.critical(self, "Proses Gagal", error_message)
