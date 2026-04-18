"""
Program Pengambilan Data RSSI UHF-RFID
Penelitian: SVR-LANDMARC Adaptif
Reader   : EL-UHF-RC4-2 (Electron Indonesia)
Protokol : Binary frame, HEAD=0xCF, 115200 8N1

Format RSSI (dokumentasi resmi electron.id):
  - RSSI = 2 Bytes, two's complement, unit = 0.1 dBm
  - Contoh: 0xFE48 = -440 (raw) = -44.0 dBm

Cara pakai:
  python rfid_data_collector.py

Library yang dibutuhkan:
  pip install pyserial openpyxl
"""

import serial
import time
import openpyxl
from openpyxl import Workbook
from datetime import datetime

# ─────────────────────────────────────────────
#  KONFIGURASI — SESUAIKAN DI SINI
# ─────────────────────────────────────────────
PORT           = "COM11"    # Ganti dengan port COM kamu (mis. COM4, COM5)
BAUDRATE       = 115200
TARGET_SAMPLES = 100       # Jumlah sampel per sesi
READER_ADDR    = 0x00      # Alamat default reader

# ─────────────────────────────────────────────
#  KONSTANTA PROTOKOL EL-UHF-RC4
# ─────────────────────────────────────────────
HEADER = 0xCF
CMD_INVENTORY_STOP = bytearray([0x00, 0x02])
CMD_INIT           = bytearray([0x00, 0x50])

# Frame inventory start lengkap dari dokumentasi resmi
# CF FF 00 01 05 00 00 00 00 00 F5 B5
FRAME_INVENTORY_START = bytes.fromhex("CFFF0001050000000000F5B5")


# ─────────────────────────────────────────────
#  CRC16 DAN FRAME BUILDER
# ─────────────────────────────────────────────
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


# ─────────────────────────────────────────────
#  KONVERSI RSSI (PERBAIKAN UTAMA)
# ─────────────────────────────────────────────
def parse_rssi(byte_high: int, byte_low: int) -> float:
    """
    Konversi 2 byte RSSI ke dBm.

    Format dari dokumentasi resmi EL-UHF-RC4:
      - 2 bytes, two's complement signed 16-bit
      - Unit: 0.1 dBm
      - Contoh: 0xFE48 = -440 = -44.0 dBm

    Sebelumnya program membaca 1 byte signed → menghasilkan nilai -3 (salah).
    Sekarang membaca 2 byte signed → akan menghasilkan nilai sekitar -44 dBm (benar).
    """
    raw_uint16 = (byte_high << 8) | byte_low
    # Konversi two's complement ke signed
    raw_signed = raw_uint16 - 0x10000 if raw_uint16 >= 0x8000 else raw_uint16
    return raw_signed / 10.0


# ─────────────────────────────────────────────
#  PARSER RESPONSE FRAME
# ─────────────────────────────────────────────
def parse_inventory_response(data: bytes) -> list:
    """
    Parsing semua inventory response frame dari buffer serial.

    Format response (CMD=0x0001, STATUS=0x00):
      HEAD | ADDR | CMD(2) | LEN | STATUS | InventoryParam | CRC(2)

    InventoryParam:
      RSSI(2) | Antenna(1) | Channel(1) | TagInfoLen(1) | TagInfo(N bytes)
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
        total    = 5 + length + 2  # header(5) + payload(length) + crc(2)

        if i + total > len(data):
            break

        status = data[i+5]
        param  = data[i+6 : i+5+length]

        if cmd_word == 0x0001 and status == 0x00 and len(param) >= 5:
            rssi_dbm     = parse_rssi(param[0], param[1])
            tag_info_len = param[4]
            tag_info     = param[5 : 5 + tag_info_len]
            epc_hex      = ' '.join(f'{b:02X}' for b in tag_info)

            tags.append({
                'epc'      : epc_hex,
                'rssi_dbm' : rssi_dbm,
                'rssi_raw' : f"0x{param[0]:02X}{param[1]:02X}",
            })

        i += total

    return tags


# ─────────────────────────────────────────────
#  KELAS RFID COLLECTOR
# ─────────────────────────────────────────────
class RFIDCollector:
    def __init__(self, port, baudrate=115200):
        self.port     = port
        self.baudrate = baudrate
        self.ser      = None

    def connect(self) -> bool:
        try:
            self.ser = serial.Serial(
                port=self.port, baudrate=self.baudrate,
                bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE, timeout=1.0
            )
            time.sleep(0.5)
            self.ser.flushInput()
            print(f"[OK] Terhubung ke reader di {self.port}")
            return True
        except serial.SerialException as e:
            print(f"[ERROR] Gagal membuka {self.port}: {e}")
            return False

    def init_reader(self):
        self.ser.write(build_command(CMD_INIT))
        time.sleep(0.3)
        resp = self.ser.read(self.ser.in_waiting or 10)
        if resp and len(resp) >= 6 and resp[0] == 0xCF:
            print(f"[OK] Reader diinisialisasi. Status byte: 0x{resp[5]:02X}")
        else:
            print("[INFO] Tidak ada respons inisialisasi.")
        self.ser.flushInput()

    def start_inventory(self):
        self.ser.write(FRAME_INVENTORY_START)
        time.sleep(0.05)

    def stop_inventory(self):
        self.ser.write(build_command(CMD_INVENTORY_STOP))
        time.sleep(0.2)
        self.ser.flushInput()

    def read_buffer(self, timeout_sec=0.4) -> bytes:
        buf   = bytearray()
        start = time.time()
        while (time.time() - start) < timeout_sec:
            n = self.ser.in_waiting
            if n:
                buf.extend(self.ser.read(n))
            else:
                time.sleep(0.01)
        return bytes(buf)

    def disconnect(self):
        if self.ser and self.ser.is_open:
            try:
                self.stop_inventory()
            except Exception:
                pass
            self.ser.close()
            print("[OK] Koneksi ditutup.")


# ─────────────────────────────────────────────
#  MODE DIAGNOSTIK
# ─────────────────────────────────────────────
def run_diagnostic(collector: RFIDCollector):
    """
    Tampilkan raw bytes dan hasil parsing RSSI secara real-time.
    Gunakan ini untuk memverifikasi nilai RSSI sudah benar sebelum
    melakukan pengambilan data sesungguhnya.
    """
    print("\n" + "="*55)
    print("  MODE DIAGNOSTIK")
    print("  Dekatkan tag ke reader. Tekan Ctrl+C untuk keluar.")
    print("  RSSI yang wajar pada jarak 1m: sekitar -40 s/d -55 dBm")
    print("="*55 + "\n")
    collector.start_inventory()
    try:
        while True:
            raw  = collector.read_buffer(1.0)
            tags = parse_inventory_response(raw)
            if tags:
                for t in tags:
                    print(f"  EPC: {t['epc']:30s} | "
                          f"RSSI raw: {t['rssi_raw']} | "
                          f"RSSI: {t['rssi_dbm']:7.1f} dBm")
            elif raw:
                hex_str = ' '.join(f'{b:02X}' for b in raw[:30])
                print(f"  [raw tidak terparsing] {hex_str}...")
    except KeyboardInterrupt:
        print("\n[INFO] Diagnostik dihentikan.")
    finally:
        collector.stop_inventory()


# ─────────────────────────────────────────────
#  PENGAMBILAN SAMPEL
# ─────────────────────────────────────────────
def collect_samples(collector, n_samples, label, jarak_cm, kondisi, skenario):
    records    = []
    target_epc = None

    print(f"\n{'='*55}")
    print(f"  Tag: {label}  |  Jarak: {jarak_cm} cm")
    print(f"  Kondisi: {kondisi}  |  Skenario: {skenario}")
    print(f"  Target: {n_samples} sampel")
    print(f"{'='*55}")
    input("  Pastikan tag sudah di posisi yang benar, lalu tekan ENTER...\n")

    collector.start_inventory()

    while len(records) < n_samples:
        raw  = collector.read_buffer(0.3)
        tags = parse_inventory_response(raw)

        for tag in tags:
            epc = tag['epc']
            if not epc:
                continue

            if target_epc is None:
                target_epc = epc
                print(f"  [INFO] Tag terkunci — EPC: {epc}")
                print(f"         RSSI pertama: {tag['rssi_dbm']:.1f} dBm "
                      f"(raw: {tag['rssi_raw']})\n")

            if epc == target_epc:
                records.append({
                    'epc'       : epc,
                    'rssi_dbm'  : tag['rssi_dbm'],
                    'timestamp' : datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                })
                count = len(records)
                if count % 50 == 0 or count == n_samples:
                    rssi_vals = [r['rssi_dbm'] for r in records]
                    avg = sum(rssi_vals) / len(rssi_vals)
                    print(f"  [{count:>3}/{n_samples}]  avg RSSI: {avg:.1f} dBm")

            if len(records) >= n_samples:
                break

    collector.stop_inventory()
    rssi_vals = [r['rssi_dbm'] for r in records]
    print(f"\n  [SELESAI] {len(records)} sampel.")
    print(f"  min={min(rssi_vals):.1f}  max={max(rssi_vals):.1f}  "
          f"avg={sum(rssi_vals)/len(rssi_vals):.1f} dBm")
    return records


# ─────────────────────────────────────────────
#  SIMPAN KE EXCEL
# ─────────────────────────────────────────────
def save_to_excel(records, label, jarak_cm, kondisi, skenario):
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"data_{skenario}_{kondisi}_{label}_{int(jarak_cm)}cm_{ts}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Data RSSI"
    ws.append(["No", "Timestamp", "EPC", "RSSI (dBm)",
               "Jarak Aktual (cm)", "Label Tag", "Kondisi", "Skenario"])
    for i, r in enumerate(records, 1):
        ws.append([i, r['timestamp'], r['epc'], r['rssi_dbm'],
                   jarak_cm, label, kondisi, skenario])
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(w + 4, 40)
    wb.save(filename)
    print(f"[OK] Tersimpan: {filename}")
    return filename


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  Program Pengambilan Data RSSI UHF-RFID")
    print("  Penelitian: SVR-LANDMARC Adaptif — Wiranto")
    print("=" * 55)
    print("\nMode:  1 = Pengambilan data  |  2 = Diagnostik (debug)")
    mode = input("Pilihan (1/2): ").strip()

    collector = RFIDCollector(PORT, BAUDRATE)
    if not collector.connect():
        print(f"Ubah variabel PORT di baris 30 script ini.")
        return
    collector.init_reader()

    if mode == "2":
        run_diagnostic(collector)
        collector.disconnect()
        return

    # Mode pengambilan data normal
    skenario  = "Skenario" + input("\nSkenario (1=2ref / 2=3ref): ").strip()
    kondisi_m = {'1':'LOS','2':'NLOS_Sedikit','3':'NLOS_Sedang','4':'NLOS_Banyak'}
    print("Kondisi: 1=LOS  2=NLOS_Sedikit  3=NLOS_Sedang  4=NLOS_Banyak")
    kondisi   = kondisi_m.get(input("Pilih (1-4): ").strip(), 'LOS')
    tag_m     = {'1':'R1','2':'R2','3':'R3','4':'Target'}
    print("Tag: 1=R1  2=R2  3=R3  4=Target")
    label_tag = tag_m.get(input("Pilih (1-4): ").strip(), 'Tag')
    jarak_cm  = float(input(f"Jarak aktual {label_tag} dari reader (cm): ").strip())

    records = []
    try:
        records = collect_samples(collector, TARGET_SAMPLES,
                                  label_tag, jarak_cm, kondisi, skenario)
        save_to_excel(records, label_tag, jarak_cm, kondisi, skenario)
        print("\n[SELESAI] Jalankan lagi untuk tag/posisi berikutnya.")
    except KeyboardInterrupt:
        print("\n[INFO] Dihentikan pengguna.")
        if records:
            save_to_excel(records, label_tag, jarak_cm, kondisi, skenario)
    finally:
        collector.disconnect()


if __name__ == "__main__":
    main()