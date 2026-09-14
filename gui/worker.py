"""
Worker - membungkus ReportProcessor di dalam QThread supaya GUI tidak
freeze selama proses berjalan (yang bisa memakan waktu lama untuk file
besar / banyak ZIP).

File ini TIDAK berisi logika bisnis apa pun - murni jembatan antara
backend (ReportProcessor) dan GUI (sinyal Qt).
"""

from pathlib import Path

from PySide6.QtCore import QThread, Signal

from backend import ReportProcessor


class ProcessWorker(QThread):

    log_line = Signal(str)
    zip_progress = Signal(int, int, str)      # index, total, nama_zip
    file_progress = Signal(int, int)          # index, total (dalam 1 zip)
    finished_ok = Signal(str)                 # path output
    stopped = Signal(str)                     # path output (dihentikan user, sebagian tersimpan)
    failed = Signal(str)                      # pesan error

    def __init__(self, zip_paths, master_file, output_folder, parent=None):
        super().__init__(parent)
        self.zip_paths = zip_paths
        self.master_file = master_file
        self.output_folder = output_folder
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def _should_stop(self):
        return self._stop_requested

    def run(self):
        try:
            processor = ReportProcessor(
                master_file=self.master_file,
                output_folder=self.output_folder,
                log=self._on_log,
            )

            output_path = processor.run(
                self.zip_paths,
                on_zip_progress=self._on_zip_progress,
                on_file_progress=self._on_file_progress,
                should_stop=self._should_stop,
            )

            if self._stop_requested:
                self.stopped.emit(str(output_path))
            else:
                self.finished_ok.emit(str(output_path))

        except Exception as e:
            self.failed.emit(str(e))

    # ---- callback yang diteruskan ke backend ----

    def _on_log(self, message):
        self.log_line.emit(str(message))

    def _on_zip_progress(self, index, total, nama_zip):
        self.zip_progress.emit(index, total, nama_zip)

    def _on_file_progress(self, index, total):
        self.file_progress.emit(index, total)
