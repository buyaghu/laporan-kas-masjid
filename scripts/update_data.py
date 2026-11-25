#!/usr/bin/env python3
"""Ambil data dashboard dari Excel (sheet 'Data Dashboard Kas', A84:A97),
validasi, lalu ganti blok DATA-START..DATA-END di index.html.
Keluar dengan kode != 0 (tanpa mengubah file) jika validasi gagal."""
import json, re, sys
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook

XLSX, HTML = sys.argv[1], sys.argv[2]
SHEET, FIRST, LAST = "Data Dashboard Kas", 84, 97

def fail(msg):
    print("GAGAL: " + msg); sys.exit(1)

wb = load_workbook(XLSX, data_only=True, read_only=True)
if SHEET not in wb.sheetnames:
    fail(f"sheet '{SHEET}' tidak ditemukan")
ws = wb[SHEET]
cells = [ws[f"A{r}"].value for r in range(FIRST, LAST + 1)]
if any(not isinstance(c, str) for c in (cells[0], *cells[2:])):
    fail("sel A84/A86:A97 kosong atau bukan teks (file belum disimpan oleh Excel / rumus belum terhitung?)")

head = cells[0].strip()
if not (head.startswith("const DATA = {") and head.endswith("};")):
    fail("A84 tidak berformat 'const DATA = {...};'")
try:
    data = json.loads(head[len("const DATA = "):-1])
except Exception as e:
    fail(f"A84 bukan JSON valid: {e}")

trx = []
for i, c in enumerate(cells[2:], start=FIRST + 2):
    c = c.strip()
    m = re.fullmatch(r"TRX\.push\((.*)\);", c, re.S)
    if not m:
        fail(f"A{i} tidak berformat 'TRX.push(...);'")
    if m.group(1).strip():
        try:
            trx += json.loads("[" + m.group(1) + "]")
        except Exception as e:
            fail(f"A{i} bukan JSON valid: {e}")

# --- validasi ---
bulan = data["bulan"]
if not bulan: fail("daftar bulan kosong")
if not trx: fail("tidak ada transaksi")
for t in trx:
    if len(t) != 8: fail(f"baris transaksi tidak 8 kolom: {t}")
prev = None
for i, b in enumerate(bulan, start=1):
    if abs(b["ko"] + b["ks"] + b["kd"] - b["ak"]) > 2:
        fail(f"{b['m']}: tiga kas tidak sama dengan saldo akhir")
    if prev is not None and abs(b["sa"] - prev) > 0:
        fail(f"{b['m']}: saldo awal bukan saldo akhir bulan sebelumnya")
    prev = b["ak"]
    rows = [t for t in trx if t[0] == i]
    tin, tout = sum(t[6] for t in rows), sum(t[7] for t in rows)
    if abs(tin - b["in"]) > 2 or abs(tout - b["out"]) > 2:
        fail(f"{b['m']}: jumlah transaksi ({tin}/{tout}) tidak cocok dengan ringkasan ({b['in']}/{b['out']})")

# --- tulis ke HTML ---
new_block = "\n".join([head, "const TRX = [];"] + [c.strip() for c in cells[2:]])
html = open(HTML, encoding="utf-8").read()
pat = re.compile(r"(// DATA-START\n).*?(\n// DATA-END)", re.S)
if not pat.search(html):
    fail("penanda DATA-START/DATA-END tidak ditemukan di index.html")
out = pat.sub(lambda m: m.group(1) + new_block + m.group(2), html, count=1)
if out == html:
    print("Tidak ada perubahan data."); sys.exit(0)
BULAN = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"]
now = datetime.now(timezone(timedelta(hours=7)))  # WIB
tgl = f"{now.day} {BULAN[now.month-1]} {now.year}"
out = re.sub(r'const TGL_INPUT = "[^"]*";', f'const TGL_INPUT = "{tgl}";', out, count=1)
open(HTML, "w", encoding="utf-8").write(out)
print(f"OK: data diperbarui ({data['update']}, {len(bulan)} bulan, {len(trx)} transaksi)")
