"""
Program Pengambilan Data RSSI UHF-RFID
Penelitian: SVR-LANDMARC Adaptif — Wiranto (NIM 11221030)
Reader   : EL-UHF-RC4-2 (Electron Indonesia)
Protokol : Binary frame, HEAD=0xCF, 115200 8N1

CARA KERJA YANG BENAR:
  - Semua tag referensi (R1, R2, dst) diambil BERSAMAAN dalam 1 sesi
  - Reader scan semua tag, data dipisah otomatis berdasarkan EPC
  - Hasil disimpan ke 1 file Excel dengan sheet terpisah per tag

SEBELUM MENJALANKAN:
  1. pip install pyserial openpyxl
  2. Ubah PORT di bawah sesuai port COM kamu
  3. Daftarkan EPC setiap tag di bagian KONFIGURASI TAG
"""

import serial
import time
from openpyxl import Workbook
from datetime import datetime

# ══════════════════════════════════════════════════════
#  KONFIGURASI UTAMA — SESUAIKAN SEBELUM DIJALANKAN
# ══════════════════════════════════════════════════════

PORT           = "COM11"   # Port COM reader kamu, mis. COM4, COM5
BAUDRATE       = 115200
TARGET_SAMPLES = 500      # Jumlah sampel per tag

# Daftarkan EPC dan jarak tiap tag di sini.
# EPC bisa dilihat dari output diagnostik (mode 2).
# Jalankan mode diagnostik dulu untuk mendapatkan EPC setiap tag.
#
# Format: "NAMA_TAG": {"epc": "XX XX XX ...", "jarak_cm": angka}
TAG_CONFIG = {
    "R1": {"epc": "E2 80 69 95 00 00 50 17 60 37 C0 F7", "jarak_cm": 100},
    "R2": {"epc": "E2 80 69 95 00 00 40 17 60 37 A8 F7", "jarak_cm": 300},
    # Untuk Skenario 2 (3 tag referensi), tambahkan R3:
    # "R3": {"epc": "XX XX XX XX XX XX XX XX XX XX XX XX", "jarak_cm": 200},
}

# ══════════════════════════════════════════════════════
#  PROTOKOL EL-UHF-RC4
# ══════════════════════════════════════════════════════

HEADER = 0xCF
READER_ADDR = 0x00
FRAME_INVENTORY_START = bytes.fromhex("CFFF0001050000000000F5B5")
CMD_INVENTORY_STOP    = bytearray([0x00, 0x02])
CMD_INIT              = bytearray([0x00, 0x50])


def calculate_crc16(data: bytes) -> bytearray:
    value = 0xFFFF
    for d in data:
        value ^= d
        for _ in range(8):
            value = (value >> 1) ^ 0x8408 if value & 0x0001 else (value >> 1)
    return bytearray([value >> 8, value & 0xFF])


def build_command(cmd: bytearray, param: bytearray = bytearray()) -> bytes:
    frame = bytearray([HEADER, READER_ADDR])
    frame.extend(cmd)
    frame.extend(bytearray([len(param)]))
    frame.extend(param)
    frame.extend(calculate_crc16(frame))
    return bytes(frame)


def parse_rssi(byte_high: int, byte_low: int) -> float:
    """2-byte two's complement, unit 0.1 dBm → dBm float."""
    raw = (byte_high << 8) | byte_low
    signed = raw - 0x10000 if raw >= 0x8000 else raw
    return signed / 10.0


def parse_inventory_response(data: bytes) -> list:
    """
    Parse semua inventory response frame dari buffer.
    Return: list of {'epc': str, 'rssi_dbm': float}
    """
    tags = []
    i = 0
    while i < len(data):
        if data[i] != 0xCF:
            i += 1
            continue
        if i + 6 >= len(data):
            break

        cmd_word = (data[i+2] << 8) | data[i+3]
        length   = data[i+4]
        total    = 5 + length + 2

        if i + total > len(data):
            break

        status = data[i+5]
        param  = data[i+6 : i+5+length]

        if cmd_word == 0x0001 and status == 0x00 and len(param) >= 5:
            rssi_dbm     = parse_rssi(param[0], param[1])
            tag_info_len = param[4]
            tag_info     = param[5 : 5 + tag_info_len]
            epc_hex      = ' '.join(f'{b:02X}' for b in tag_info)
            if epc_hex:
                tags.append({'epc': epc_hex, 'rssi_dbm': rssi_dbm})

        i += total
    return tags


# ══════════════════════════════════════════════════════
#  KELAS READER
# ══════════════════════════════════════════════════════

class RFIDReader:
    def __init__(self, port, baudrate=115200):
        self.ser = None
        self.port = port
        self.baudrate = baudrate

    def connect(self) -> bool:
        try:
            self.ser = serial.Serial(
                port=self.port, baudrate=self.baudrate,
                bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE, timeout=1.0
            )
            time.sleep(0.5)
            self.ser.flushInput()
            print(f"[OK] Terhubung ke {self.port}")
            return True
        except serial.SerialException as e:
            print(f"[ERROR] {e}")
            return False

    def init(self):
        self.ser.write(build_command(CMD_INIT))
        time.sleep(0.3)
        self.ser.flushInput()
        print("[OK] Reader diinisialisasi.")

    def start(self):
        self.ser.write(FRAME_INVENTORY_START)
        time.sleep(0.05)

    def stop(self):
        self.ser.write(build_command(CMD_INVENTORY_STOP))
        time.sleep(0.2)
        self.ser.flushInput()

    def read_buffer(self, timeout=0.3) -> bytes:
        buf = bytearray()
        t0  = time.time()
        while (time.time() - t0) < timeout:
            n = self.ser.in_waiting
            if n:
                buf.extend(self.ser.read(n))
            else:
                time.sleep(0.005)
        return bytes(buf)

    def disconnect(self):
        if self.ser and self.ser.is_open:
            try:
                self.stop()
            except Exception:
                pass
            self.ser.close()
            print("[OK] Koneksi ditutup.")


# ══════════════════════════════════════════════════════
#  MODE DIAGNOSTIK — scan EPC semua tag yang terbaca
# ══════════════════════════════════════════════════════

def run_diagnostic(reader: RFIDReader):
    """
    Gunakan ini PERTAMA KALI untuk mengetahui EPC setiap tag.
    Taruh tag satu per satu di depan reader, catat EPC-nya,
    lalu isi TAG_CONFIG di bagian atas script.
    """
    print("\n" + "="*60)
    print("  MODE DIAGNOSTIK — Scan EPC & Cek Nilai RSSI")
    print("  Taruh tag di depan reader.")
    print("  RSSI normal 1m: -40 s/d -55 dBm")
    print("  RSSI normal 3m: -65 s/d -80 dBm")
    print("  Tekan Ctrl+C untuk keluar.")
    print("="*60 + "\n")

    seen_epcs = {}
    reader.start()
    try:
        while True:
            raw  = reader.read_buffer(1.0)
            tags = parse_inventory_response(raw)
            for t in tags:
                epc  = t['epc']
                rssi = t['rssi_dbm']
                if epc not in seen_epcs:
                    seen_epcs[epc] = []
                seen_epcs[epc].append(rssi)
                avg = sum(seen_epcs[epc]) / len(seen_epcs[epc])
                print(f"  EPC: {epc}  |  RSSI: {rssi:6.1f} dBm  (avg: {avg:.1f})")
    except KeyboardInterrupt:
        print("\n\n[INFO] EPC yang terdeteksi:")
        for epc, rvals in seen_epcs.items():
            avg = sum(rvals) / len(rvals)
            print(f"  {epc}  →  avg RSSI: {avg:.1f} dBm  ({len(rvals)} bacaan)")
        print("\nSalin EPC di atas ke TAG_CONFIG di bagian atas script.")
    finally:
        reader.stop()


# ══════════════════════════════════════════════════════
#  PENGAMBILAN DATA UTAMA — semua tag bersamaan
# ══════════════════════════════════════════════════════

def collect_all_tags(reader: RFIDReader, tag_config: dict,
                     kondisi: str, skenario: str) -> dict:
    """
    Kumpulkan TARGET_SAMPLES sampel untuk SETIAP tag secara bersamaan.
    Reader scan terus, setiap bacaan dicocokkan ke EPC yang terdaftar.

    Return: dict {"R1": [records], "R2": [records], ...}
    """
    # Inisialisasi tempat penyimpanan data per tag
    records = {name: [] for name in tag_config}
    # Buat lookup: EPC string → nama tag
    epc_to_name = {cfg['epc']: name for name, cfg in tag_config.items()}

    # Cek berapa tag yang sudah penuh
    def all_done():
        return all(len(records[n]) >= TARGET_SAMPLES for n in tag_config)

    def progress_str():
        parts = [f"{n}:{len(records[n])}/{TARGET_SAMPLES}" for n in tag_config]
        return "  |  ".join(parts)

    print(f"\n{'='*60}")
    print(f"  Pengambilan data: {skenario} — {kondisi}")
    print(f"  Tag terdaftar: {', '.join(tag_config.keys())}")
    print(f"  Target: {TARGET_SAMPLES} sampel per tag")
    print(f"{'='*60}")
    print("  Pastikan SEMUA tag referensi sudah di posisinya.")
    input("  Tekan ENTER untuk mulai...\n")

    reader.start()
    last_print = time.time()

    try:
        while not all_done():
            raw  = reader.read_buffer(0.2)
            tags = parse_inventory_response(raw)

            for t in tags:
                epc  = t['epc']
                name = epc_to_name.get(epc)

                if name and len(records[name]) < TARGET_SAMPLES:
                    records[name].append({
                        'epc'       : epc,
                        'rssi_dbm'  : t['rssi_dbm'],
                        'timestamp' : datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    })

            # Tampilkan progress setiap 2 detik
            if time.time() - last_print >= 2.0:
                print(f"  [{progress_str()}]")
                last_print = time.time()

    except KeyboardInterrupt:
        print("\n[INFO] Dihentikan pengguna.")

    reader.stop()

    # Ringkasan
    print(f"\n  [SELESAI]")
    for name, recs in records.items():
        if recs:
            rssi_vals = [r['rssi_dbm'] for r in recs]
            print(f"  {name} ({tag_config[name]['jarak_cm']}cm): "
                  f"{len(recs)} sampel  |  "
                  f"avg={sum(rssi_vals)/len(rssi_vals):.1f}  "
                  f"min={min(rssi_vals):.1f}  max={max(rssi_vals):.1f} dBm")

    return records


# ══════════════════════════════════════════════════════
#  SIMPAN KE EXCEL — 1 file, sheet terpisah per tag
# ══════════════════════════════════════════════════════

def save_to_excel(records: dict, tag_config: dict,
                  kondisi: str, skenario: str) -> str:
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"data_{skenario}_{kondisi}_{ts}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)  # Hapus sheet default

    for name, recs in records.items():
        if not recs:
            continue
        ws = wb.create_sheet(title=name)
        jarak = tag_config[name]['jarak_cm']

        ws.append(["No", "Timestamp", "EPC", "RSSI (dBm)",
                   "Jarak Aktual (cm)", "Label Tag", "Kondisi", "Skenario"])
        for i, r in enumerate(recs, 1):
            ws.append([i, r['timestamp'], r['epc'], r['rssi_dbm'],
                       jarak, name, kondisi, skenario])

        # Auto-width
        for col in ws.columns:
            w = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(w + 4, 40)

    # Sheet ringkasan
    ws_sum = wb.create_sheet(title="Ringkasan", index=0)
    ws_sum.append(["Tag", "Jarak (cm)", "Jumlah Sampel",
                   "RSSI Avg (dBm)", "RSSI Min (dBm)", "RSSI Max (dBm)",
                   "Kondisi", "Skenario"])
    for name, recs in records.items():
        if recs:
            rssi_vals = [r['rssi_dbm'] for r in recs]
            ws_sum.append([
                name,
                tag_config[name]['jarak_cm'],
                len(recs),
                round(sum(rssi_vals)/len(rssi_vals), 2),
                min(rssi_vals),
                max(rssi_vals),
                kondisi,
                skenario
            ])

    wb.save(filename)
    print(f"\n[OK] Tersimpan: {filename}")
    print(f"     Sheet: Ringkasan + {', '.join(records.keys())}")
    return filename


# ══════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  Program Pengambilan Data RSSI UHF-RFID")
    print("  Penelitian: SVR-LANDMARC Adaptif — Wiranto")
    print("=" * 60)

    print("\nMode:")
    print("  1 = Pengambilan data (semua tag bersamaan)")
    print("  2 = Diagnostik (cari EPC tag & cek nilai RSSI)")
    mode = input("Pilihan (1/2): ").strip()

    reader = RFIDReader(PORT, BAUDRATE)
    if not reader.connect():
        print(f"\nUbah variabel PORT di baris 35 script ini.")
        return
    reader.init()

    if mode == "2":
        run_diagnostic(reader)
        reader.disconnect()
        return

    # ── Mode pengambilan data ──────────────────────────
    print("\nKonfigurasi sesi:")
    skenario  = "Skenario" + input("  Skenario (1=2ref / 2=3ref): ").strip()
    kondisi_m = {'1':'LOS','2':'NLOS_Sedikit','3':'NLOS_Sedang','4':'NLOS_Banyak'}
    print("  Kondisi: 1=LOS  2=NLOS_Sedikit  3=NLOS_Sedang  4=NLOS_Banyak")
    kondisi = kondisi_m.get(input("  Pilih kondisi (1-4): ").strip(), 'LOS')

    # Filter TAG_CONFIG sesuai skenario
    if "1" in skenario:
        active_tags = {k: v for k, v in TAG_CONFIG.items() if k in ('R1', 'R2')}
    else:
        active_tags = TAG_CONFIG.copy()

    print(f"\n  Tag aktif: {list(active_tags.keys())}")
    for name, cfg in active_tags.items():
        print(f"    {name}: jarak {cfg['jarak_cm']} cm | EPC: {cfg['epc']}")

    confirm = input("\n  Sudah benar? (y/n): ").strip().lower()
    if confirm != 'y':
        print("  Batalkan. Edit TAG_CONFIG di bagian atas script.")
        reader.disconnect()
        return

    records = {}
    try:
        records = collect_all_tags(reader, active_tags, kondisi, skenario)
        save_to_excel(records, active_tags, kondisi, skenario)
        print("\n[SELESAI] Jalankan lagi untuk sesi berikutnya.")
    except Exception as e:
        print(f"\n[ERROR] {e}")
        if records:
            print("Menyimpan data yang sudah terkumpul...")
            save_to_excel(records, active_tags, kondisi, skenario)
    finally:
        reader.disconnect()


if __name__ == "__main__":
    main()