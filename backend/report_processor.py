"""
ReportProcessor - backend inti aplikasi Laporan Harian Otomatis.

PENTING: Seluruh logika di file ini (mapping header, deteksi baris data,
anti-duplikat, penggabungan ke master) ADALAH SALINAN PERSIS dari
`main.py` versi asli. TIDAK ADA perubahan algoritma di sini.

Perubahan yang dilakukan HANYA bersifat struktural, supaya class ini bisa
dipakai ulang (reusable) oleh GUI maupun CLI:

1. Path (INPUT_FOLDER / MASTER_FILE / OUTPUT_FILE) yang sebelumnya
   hardcoded sebagai class attribute, sekarang diterima sebagai parameter
   di __init__ / run().
2. Semua pemanggilan print(...) diganti menjadi self.log(...), yaitu
   sebuah callback (default-nya tetap print, jadi perilaku CLI lama tidak
   berubah). Ini supaya GUI bisa menampilkan log secara realtime.
3. Import win32com.client dibungkus try/except, karena modul ini hanya
   tersedia di Windows. Kalau tidak tersedia (mis. di macOS), fitur
   konversi .xls lama akan melempar error yang jelas saat dipanggil,
   BUKAN meng-crash seluruh aplikasi saat import.
4. Ditambahkan hook progress opsional (on_zip_progress / on_file_progress)
   di process_zip/run, murni untuk keperluan progress bar di GUI - tidak
   mengubah urutan maupun cara pemrosesan data sama sekali.
"""

from pathlib import Path
from zipfile import ZipFile
from tempfile import TemporaryDirectory
from datetime import datetime, date

from openpyxl import load_workbook

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    HAS_WIN32COM = False


class ReportProcessor:

    SHEETS = [
        "LSDK",
        "LSTAKDKP",
        "LTAKDK",
        "LRTAK"
    ]

    # Kolom yang dipakai sebagai identitas unik tiap baris (untuk anti-duplikat).
    # None artinya sheet ini tidak punya kolom identitas per-baris di master,
    # jadi pengecekan duplikat dilakukan per-hari (lihat DAY_LEVEL_SHEETS).
    IDENTITY_COLUMN = {
        "LSDK": None,
        "LSTAKDKP": "Kode Aset Keuangan Digital",
        "LTAKDK": "Uraian",
        "LRTAK": "Kode Aset Keuangan Digital",
    }

    # Kolom "jangkar" yang menentukan apakah suatu baris masih berisi data
    # asli atau sudah masuk baris penutup (Total / "+ Baris" / "<EOR>" / kosong).
    # Kolom pertama (A) di source sering berisi teks marker yang tidak relevan
    # (mis. "+ Baris", "<EOR>"), jadi kita cek kolom identitas asli, bukan
    # kolom pertama yang kebetulan terisi.
    STOP_ANCHOR_COLUMN = {
        "LSDK": "Sesi",
        "LSTAKDKP": "Kode Aset Keuangan Digital",
        "LTAKDK": "Uraian",
        "LRTAK": "Kode Aset Keuangan Digital",
    }

    # Sheet yang mastenya TIDAK menyimpan kolom identitas per-baris
    # (mis. "Kode Aset Keuangan Digital"), jadi anti-duplikat dilakukan
    # per-hari saja: kalau tanggal itu utk PT ini sudah pernah masuk,
    # seluruh baris hari itu dilewati.
    DAY_LEVEL_SHEETS = {"LSDK", "LRTAK"}

    def __init__(self, master_file, output_folder, log=print):
        """
        master_file   : path ke file Template Excel (hasil "Browse Template Excel" di GUI)
        output_folder : folder tujuan output (hasil "Browse Folder Output" di GUI)
        log           : callback(str) untuk menampilkan log. Default: print (perilaku CLI lama).
        """

        self.master_file = Path(master_file)
        self.output_folder = Path(output_folder)
        self.log = log

        self.master = load_workbook(self.master_file)

        # cache anti-duplikat di RAM, per sheet
        self.existing_keys = {sheet: set() for sheet in self.SHEETS}

        self._load_existing_keys()

    # =====================================
    # HELPER
    # =====================================

    def clean(self, value):

        if value is None:
            return ""

        return (
            str(value)
            .replace('="', "")
            .replace('"', "")
            .replace("\n", " ")
            .replace("\r", " ")
            .strip()
        )

    def normalize_periode(self, value):
        """Samakan format periode (datetime/date/string) supaya bisa dibandingkan sebagai key."""

        if isinstance(value, (datetime, date)):
            return value.strftime("%Y-%m-%d")

        return self.clean(value)

    # =====================================
    # CONVERT XLS -> XLSX
    # =====================================
    def convert_xls_to_xlsx(self, xls_path):

        if not HAS_WIN32COM:
            raise RuntimeError(
                "File Excel lama (.xls) terdeteksi, tapi konversi otomatis "
                "hanya didukung di Windows (butuh Microsoft Excel ter-install). "
                f"File dilewati: {xls_path.name}"
            )

        xlsx_path = xls_path.with_name(xls_path.stem + "_converted.xlsx")

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            self.log(f"OPEN : {xls_path}")
            wb = excel.Workbooks.Open(str(xls_path.resolve()))

            self.log(f"SAVE : {xlsx_path}")
            wb.SaveAs(str(xlsx_path.resolve()), FileFormat=51)

            wb.Close(False)

        finally:
            excel.Quit()

        self.log(f"EXISTS : {xlsx_path.exists()}")

        return xlsx_path

    # =====================================
    # DATA UMUM
    # =====================================

    def read_data_umum(self, wb):

        ws = wb["Data Umum"]

        return {
            "periode": ws["D11"].value,
            "kode": ws["D6"].value,
            "nama": ws["D7"].value
        }

    # =====================================
    # MASTER HEADER
    # =====================================

    def build_master_header(self, ws, sheet_name):

        header = {}

        # LTAKDK di master juga punya header 2 baris (row 1 = parent, row 2 = child)
        if sheet_name == "LTAKDK":

            parent = ""
            remaining = 0

            for col in range(1, ws.max_column + 1):

                top = self.clean(ws.cell(row=1, column=col).value)
                bottom = self.clean(ws.cell(row=2, column=col).value)

                if top != "":
                    parent = top
                    remaining = 2  # tiap parent "Pihak ..." cuma menaungi 2 kolom anak

                if parent in [
                    "Pihak Terafiliasi",
                    "Pihak Tidak Terafiliasi"
                ] and remaining > 0:

                    if bottom != "":
                        header[f"{parent} {bottom}"] = col

                    remaining -= 1

                else:

                    label = top if top != "" else bottom

                    if label != "":
                        header[label] = col

            return header

        for col in range(1, ws.max_column + 1):

            value = self.clean(
                ws.cell(row=1, column=col).value
            )

            if value != "":
                header[value] = col

        return header

    # =====================================
    # SOURCE HEADER
    # =====================================

    def build_source_header(self, ws, sheet_name):

        header = {}

        # ==========================
        # LTAKDK (2 baris header)
        # ==========================
        if sheet_name == "LTAKDK":

            parent = ""
            remaining = 0

            for col in range(1, ws.max_column + 1):

                top = self.clean(ws.cell(row=14, column=col).value)
                bottom = self.clean(ws.cell(row=15, column=col).value)

                if top != "":
                    parent = top
                    remaining = 2

                if parent in [
                    "Pihak Terafiliasi",
                    "Pihak Tidak Terafiliasi"
                ] and remaining > 0:

                    if bottom != "":
                        header[f"{parent} {bottom}"] = col

                    remaining -= 1

                else:

                    if top != "":
                        header[top] = col
                    elif bottom != "":
                        header[bottom] = col

            return header

        # ==========================
        # SHEET LAIN
        # ==========================

        header_row = 14

        for col in range(1, ws.max_column + 1):

            value = self.clean(
                ws.cell(row=header_row, column=col).value
            )

            if value != "":
                header[value] = col

        return header

    # =====================================
    # BUILD MAPPING
    # =====================================

    def build_mapping(self, source_header, master_header):

        mapping = {}

        ignore = {
            "Periode",
            "Kode Perusahaan",
            "Nama Perusahaan",
            "formula"
        }

        for name, target_col in master_header.items():

            if name in ignore:
                continue

            if name not in source_header:
                continue

            mapping[source_header[name]] = target_col

        return mapping

    # =====================================
    # FIND DATA START ROW
    # =====================================

    def find_data_start_row(self, ws):

        for row in range(1, ws.max_row + 1):

            for col in range(1, ws.max_column + 1):

                value = self.clean(
                    ws.cell(row=row, column=col).value
                )

                if value.lower() == "nomor baris":
                    return row + 2

        return 16

    # =====================================
    # LOAD EXISTING KEYS (CACHE ANTI-DUPLIKAT)
    # =====================================

    def _load_existing_keys(self):

        for sheet_name in self.SHEETS:

            ws = self.master[sheet_name]

            master_header = self.build_master_header(ws, sheet_name)

            periode_col = master_header.get("Periode")
            kode_col = master_header.get("Kode Perusahaan")

            identity_name = self.IDENTITY_COLUMN[sheet_name]
            identity_col = (
                master_header.get(identity_name)
                if identity_name else None
            )

            # LTAKDK datanya mulai baris 3 (row 1-2 = header)
            start_row = 3 if sheet_name == "LTAKDK" else 2

            for row in range(start_row, ws.max_row + 1):

                kode = (
                    ws.cell(row=row, column=kode_col).value
                    if kode_col else None
                )

                periode = (
                    ws.cell(row=row, column=periode_col).value
                    if periode_col else None
                )

                if kode in (None, "") and periode in (None, ""):
                    continue

                kode_key = self.clean(kode)
                periode_key = self.normalize_periode(periode)

                if sheet_name in self.DAY_LEVEL_SHEETS:

                    key = (kode_key, periode_key)

                else:

                    identity_val = (
                        self.clean(ws.cell(row=row, column=identity_col).value)
                        if identity_col else ""
                    )

                    key = (kode_key, periode_key, identity_val)

                self.existing_keys[sheet_name].add(key)

    # =====================================
    # PROCESS SATU SHEET
    # =====================================

    def process_sheet(self, info, source, target, sheet_name):

        master_header = self.build_master_header(target, sheet_name)
        source_header = self.build_source_header(source, sheet_name)

        mapping = self.build_mapping(source_header, master_header)

        identity_name = self.IDENTITY_COLUMN[sheet_name]
        identity_source_col = (
            source_header.get(identity_name)
            if identity_name else None
        )

        kode_key = self.clean(info["kode"])
        periode_key = self.normalize_periode(info["periode"])

        tanggal_str = ""
        if isinstance(info["periode"], (datetime, date)):
            tanggal_str = info["periode"].strftime("%d%m%Y")

        # ================================
        # SHEET DAY-LEVEL (LSDK)
        # kalau hari ini utk PT ini sudah pernah masuk, skip semua barisnya
        # ================================
        if sheet_name in self.DAY_LEVEL_SHEETS:

            day_key = (kode_key, periode_key)

            if day_key in self.existing_keys[sheet_name]:
                self.log(f"         (sudah ada, dilewati: {sheet_name} {periode_key})")
                return

        stop_anchor_name = self.STOP_ANCHOR_COLUMN[sheet_name]
        stop_anchor_col = source_header.get(stop_anchor_name)

        target_row = target.max_row + 1
        row = self.find_data_start_row(source)

        inserted_any = False

        while True:

            anchor_value = ""
            if stop_anchor_col:
                anchor_value = self.clean(source.cell(row=row, column=stop_anchor_col).value)

            if anchor_value == "" or anchor_value.lower() == "total":
                break

            # ================================
            # CEK DUPLIKAT PER-BARIS (sheet selain LSDK)
            # ================================
            if sheet_name not in self.DAY_LEVEL_SHEETS:

                identity_val = (
                    self.clean(source.cell(row=row, column=identity_source_col).value)
                    if identity_source_col else ""
                )

                row_key = (kode_key, periode_key, identity_val)

                if row_key in self.existing_keys[sheet_name]:
                    row += 1
                    continue

                self.existing_keys[sheet_name].add(row_key)

            target.cell(
                row=target_row,
                column=master_header["Periode"]
            ).value = info["periode"]

            target.cell(
                row=target_row,
                column=master_header["Kode Perusahaan"]
            ).value = info["kode"]

            target.cell(
                row=target_row,
                column=master_header["Nama Perusahaan"]
            ).value = info["nama"]

            if "formula" in master_header:
                target.cell(
                    row=target_row,
                    column=master_header["formula"]
                ).value = tanggal_str

            for source_col, target_col in mapping.items():

                target.cell(
                    row=target_row,
                    column=target_col
                ).value = source.cell(
                    row=row,
                    column=source_col
                ).value

            inserted_any = True

            row += 1
            target_row += 1

        if sheet_name in self.DAY_LEVEL_SHEETS and inserted_any:
            self.existing_keys[sheet_name].add((kode_key, periode_key))

    # =====================================
    # PROCESS SATU EXCEL
    # =====================================

    def process_excel(self, excel_path):

        self.log(f"   Membaca : {excel_path.name}")

        # Jika file .xls, convert dulu ke .xlsx
        # Cek signature file (bukan ekstensi)
        with open(excel_path, "rb") as f:
            header = f.read(8)

        # Excel lama (.xls) walaupun namanya .xlsx
        if header.startswith(b"\xD0\xCF\x11\xE0"):
            self.log("      File Excel lama terdeteksi, mengonversi ke .xlsx...")
            excel_path = self.convert_xls_to_xlsx(excel_path)

        wb = load_workbook(excel_path, data_only=True)

        required = [
            "Data Umum",
            "LSDK",
            "LSTAKDKP",
            "LTAKDK",
            "LRTAK"
        ]

        for sheet in required:

            if sheet not in wb.sheetnames:
                self.log(f"      Sheet {sheet} tidak ada.")
                wb.close()
                return

        info = self.read_data_umum(wb)

        for sheet_name in self.SHEETS:

            self.log(f"      Sheet : {sheet_name}")

            source = wb[sheet_name]
            target = self.master[sheet_name]

            self.process_sheet(info, source, target, sheet_name)

        wb.close()

    # =====================================
    # PROCESS SATU ZIP
    # =====================================

    def process_zip(self, zip_path, on_file_progress=None, should_stop=None):
        """
        on_file_progress: callback opsional (index, total) dipanggil tiap
        file Excel selesai diproses. Murni untuk progress bar GUI - tidak
        mengubah urutan/cara pemrosesan.
        should_stop: callback opsional () -> bool. Dicek di antara file
        (bukan di tengah satu file), supaya proses bisa dihentikan dengan
        aman tanpa merusak data yang sedang ditulis.
        """

        self.log("")
        self.log("=" * 60)
        self.log(f"Processing : {zip_path.name}")
        self.log("=" * 60)

        with TemporaryDirectory() as temp:

            with ZipFile(zip_path, "r") as zip_file:
                zip_file.extractall(temp)

            excel_files = sorted(
                list(Path(temp).rglob("*.xlsx")) +
                list(Path(temp).rglob("*.xls"))
            )

            excel_files = [
                f for f in excel_files
                if not f.name.startswith("~$")
                and "__MACOSX" not in str(f)
            ]

            total = len(excel_files)
            self.log(f"Jumlah Excel : {total}")

            for index, excel in enumerate(excel_files, start=1):

                if should_stop and should_stop():
                    self.log("")
                    self.log("Dihentikan oleh pengguna.")
                    return

                self.log(f"[{index}/{total}]")

                try:
                    self.process_excel(excel)
                except Exception as e:
                    self.log(f"ERROR : {excel.name}")
                    self.log(str(e))
                    continue
                finally:
                    if on_file_progress:
                        on_file_progress(index, total)

    # =====================================
    # SAVE
    # =====================================

    def save(self):

        self.output_folder.mkdir(parents=True, exist_ok=True)

        output = self.output_folder / (
            "Laporan Harian 2026_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )

        self.master.save(output)

        self.log("")
        self.log("=" * 60)
        self.log("MASTER BERHASIL DISIMPAN")
        self.log(str(output))
        self.log("=" * 60)

        return output

    # =====================================
    # RUN (orchestration - dulu ada di main() versi CLI)
    # =====================================

    def run(self, zip_paths, on_zip_progress=None, on_file_progress=None, should_stop=None):
        """
        zip_paths: list Path file ZIP yang akan diproses.
        on_zip_progress: callback opsional (index, total, nama_zip) dipanggil
                         tiap mulai memproses satu ZIP.
        on_file_progress: diteruskan ke process_zip (lihat di atas).
        should_stop: callback opsional () -> bool, diteruskan ke process_zip
                     dan juga dicek di antara ZIP (bukan di tengah satu ZIP).

        Return: Path file output hasil save() (tetap disimpan meski
        dihentikan lebih awal, supaya data yang sudah sempat digabung
        tidak hilang).
        """

        total_zip = len(zip_paths)

        for index, zip_file in enumerate(zip_paths, start=1):

            if should_stop and should_stop():
                self.log("")
                self.log("Dihentikan oleh pengguna.")
                break

            if on_zip_progress:
                on_zip_progress(index, total_zip, zip_file.name)

            self.process_zip(zip_file, on_file_progress=on_file_progress, should_stop=should_stop)

        return self.save()
