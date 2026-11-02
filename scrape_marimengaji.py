#!/usr/bin/env python3
"""
Scraper jadwal kajian marimengaji.com -> Excel + CSV (dijalankan tiap hari).

Alur:
  1. Buka /kajian/cari dengan browser headless, klik "Tampilkan Lebih Banyak"
     sampai habis (data dimuat via JavaScript), lalu ambil semua kartu kajian.
  2. Parse tiap kartu (judul, ustadz, hari/pekan, jam, masjid, link).
  3. Lengkapi alamat + koordinat dari masjid_master.csv (tidak dicari ulang tiap hari).
  4. Hitung 'Tanggal' = kemunculan berikutnya dari HARI INI (aturan pekan: tgl 1-7 = pekan 1, dst).
  5. Tulis output/jadwal_kajian_marimengaji.xlsx dan .csv

Pakai:  python scrape_marimengaji.py            (normal)
        python scrape_marimengaji.py --debug    (simpan HTML & kartu mentah ke output/debug/)
"""
import csv, re, sys, json, datetime as dt
from pathlib import Path

BASE = "https://marimengaji.com"
OUT = Path("output"); OUT.mkdir(exist_ok=True)
DEBUG = "--debug" in sys.argv
TODAY = dt.date.today()

HARI = {"Senin": 0, "Selasa": 1, "Rabu": 2, "Kamis": 3, "Jumat": 4, "Sabtu": 5, "Ahad": 6}
BULAN = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "Mei": 5, "Jun": 6, "Jul": 7,
         "Agu": 8, "Agt": 8, "Sep": 9, "Okt": 10, "Nov": 11, "Des": 12}
KATEGORI = {"Kajian Rutin", "Tabligh Akbar", "Khutbah Jumat", "Dauroh Ilmiyah", "Kegiatan Lainnya"}
BADGE = {"Muslimah", "Digantikan", "Menunggu", "Baru", "Diperbarui"}


# ---------------------------------------------------------------- scraping
JS_CARDS = """
() => [...document.querySelectorAll('h3 a[href*="/s/"]')].map(a => {
  let el = a.closest('h3');
  while (el.parentElement &&
         el.parentElement.querySelectorAll('h3 a[href*="/s/"]').length === 1) el = el.parentElement;
  const maps = [...el.querySelectorAll('a')].map(x => x.href).find(h => /maps|goo\\.gl/.test(h)) || '';
  const struck = !!a.querySelector('s,del,strike') ||
                 getComputedStyle(a).textDecorationLine.includes('line-through');
  return {text: el.innerText, url: a.href, title: a.innerText.trim(), struck, maps};
})
"""

def scrape_cards():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(
            locale="id-ID",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
        page = ctx.new_page()
        page.set_default_timeout(90000)
        # jangan tunggu "networkidle" (website tidak pernah benar-benar diam); cukup tunggu kartu kajian muncul
        page.goto(f"{BASE}/kajian/cari", wait_until="domcontentloaded", timeout=90000)
        try:
            page.wait_for_selector('h3 a[href*="/s/"]', timeout=90000)
        except Exception:
            print("JUDUL HALAMAN:", page.title())
            print("ISI HALAMAN (awal):", page.inner_text("body")[:600].replace("\n", " | "))
            raise
        for _ in range(60):                       # klik "Tampilkan Lebih Banyak" sampai habis
            btn = page.get_by_text(re.compile(r"Tampilkan Lebih Banyak"))
            if btn.count() == 0 or not btn.first.is_visible():
                break
            before = page.locator('h3 a[href*="/s/"]').count()
            btn.first.click()
            try:
                page.wait_for_function(
                    "n => document.querySelectorAll('h3 a[href*=\"/s/\"]').length > n",
                    arg=before, timeout=20000)
            except Exception:
                break
        cards = page.evaluate(JS_CARDS)
        if DEBUG:
            d = OUT / "debug"; d.mkdir(exist_ok=True)
            (d / "page.html").write_text(page.content(), encoding="utf-8")
            (d / "cards.json").write_text(json.dumps(cards, ensure_ascii=False, indent=1), encoding="utf-8")
        b.close()
    return cards


# ---------------------------------------------------------------- parsing
def parse_card(c):
    lines = [l.strip() for l in c["text"].splitlines() if l.strip()]
    title = c["title"]
    try:
        i = next(k for k, l in enumerate(lines) if l.replace("~", "") == title.replace("~", ""))
    except StopIteration:
        return None
    before, after = lines[:i], lines[i + 1:]
    after = [l for l in after if l not in ("Lihat Maps", "Detail Kajian", "Lihat Detail")]
    kategori = next((l for l in before if l in KATEGORI), "Kajian Rutin")
    badges = {b for b in BADGE if any(b in l.split() for l in before)}
    jam_parts = []
    for l in before:
        if l in KATEGORI:
            continue
        words = [w for w in l.split() if w not in BADGE]
        if words:
            jam_parts.append(" ".join(words))
    jam = " ".join(jam_parts)
    if len(after) < 3:
        return None
    ustadz = re.sub(r"^(Ustadz|Ustadzah)\s+", "", after[0])
    sched, place = after[1], after[2]
    hari = pekan = tgl = None
    m = re.match(r"^(Ahad|Senin|Selasa|Rabu|Kamis|Jumat|Sabtu)\s*·\s*(.*)$", sched)
    if m:
        hari, p = m.group(1), m.group(2)
        pekan = "12345" if "Tiap" in p else "".join(re.findall(r"\d", p)) or "12345"
    else:
        d = re.search(r"(\d{1,2})\s+([A-Za-z]{3})\w*\s+(\d{4})", sched)
        if d:
            tgl = dt.date(int(d.group(3)), BULAN[d.group(2).title()], int(d.group(1)))
    nama_masjid = place.rsplit(",", 1)[0].strip()
    kota_web = place.rsplit(",", 1)[1].strip() if "," in place else ""
    return dict(judul=title.replace("~", ""), ustadz=ustadz, hari=hari, pekan=pekan, tanggal_tetap=tgl,
                jam=jam, kategori=kategori, masjid=nama_masjid, kota_web=kota_web, url=c["url"],
                maps=c["maps"], coret=c["struck"], muslimah="Muslimah" in badges,
                digantikan="Digantikan" in badges)


def next_date(hari, pekan):
    for i in range(70):
        d = TODAY + dt.timedelta(days=i)
        if d.weekday() == HARI[hari] and str((d.day - 1) // 7 + 1) in pekan:
            return d


def pekan_txt(p):
    return "Tiap pekan" if p == "12345" else "Pekan ke-" + ", ".join(p)


def jam_key(j):
    m = re.match(r"(\d{1,2})[.:](\d{2})", j)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    for k, v in (("Subuh", 5), ("Dzuhur", 12), ("Ashar", 15), ("Magh", 18), ("Magr", 18)):
        if k in j:
            return v * 60
    return 999


# ---------------------------------------------------------------- build rows
def load_master():
    with open("masjid_master.csv", encoding="utf-8") as f:
        return {r["nama_masjid"].strip().lower(): r for r in csv.DictReader(f)}


RECS = []   # data ringkas untuk aplikasi web (docs/data.json)

def build_rows(cards, master):
    rows, unknown = [], set()
    for c in cards:
        k = parse_card(c)
        if not k:
            continue
        mm = master.get(k["masjid"].lower())
        if not mm:
            unknown.add(k["masjid"])
            mm = dict(alamat="", kota=k["kota_web"], latitude="", longitude="", link_maps=k["maps"])
        if k["hari"]:
            d, ulang = next_date(k["hari"], k["pekan"]), f'{k["hari"]}, {pekan_txt(k["pekan"])}'
        else:
            d, ulang = k["tanggal_tetap"], "Sekali (tanggal tetap)"
        if d is None or d < TODAY:          # acara sekali yang sudah lewat tidak dimasukkan
            continue
        if not k["coret"]:
            RECS.append([k["judul"], k["ustadz"], k["hari"] or "", k["pekan"] or "", k["jam"], k["masjid"], k["url"],
                         ("M" if k["muslimah"] else "") + ("G" if k["digantikan"] else ""),
                         k["tanggal_tetap"].isoformat() if k["tanggal_tetap"] else ""])
        cat = []
        if k["muslimah"]: cat.append("Khusus Muslimah")
        if k["digantikan"]: cat.append("Berstatus 'Digantikan' di website")
        if k["coret"]: cat.append("DICORET di website (kemungkinan tidak aktif)")
        if k["kategori"] != "Kajian Rutin": cat.append(k["kategori"])
        rows.append([d, k["judul"], k["ustadz"], k["masjid"], mm["alamat"], mm["kota"],
                     float(mm["latitude"]) if mm["latitude"] else None,
                     float(mm["longitude"]) if mm["longitude"] else None,
                     k["jam"], k["url"], ulang, "; ".join(cat), mm["link_maps"] or k["maps"], jam_key(k["jam"])])
    rows.sort(key=lambda r: (r[0], r[13]))
    return [r[:13] for r in rows], unknown


HEADERS = ["Tanggal", "Judul Kajian", "Nama Ustadz", "Nama Masjid", "Tempat (Alamat Lengkap)",
           "Nama Kota/Kabupaten", "Latitude", "Longitude", "Jam", "Sumber Data",
           "Jadwal Berulang", "Catatan", "Link Maps"]


# ---------------------------------------------------------------- output
def write_outputs(rows, unknown):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = "Jadwal Kajian"
    ws.append(HEADERS)
    thin = Side(style="thin", color="D0D7D3"); brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in range(1, len(HEADERS) + 1):
        x = ws.cell(1, c); x.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        x.fill = PatternFill("solid", fgColor="1E4D35"); x.border = brd
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows: ws.append(r)
    for ri in range(2, ws.max_row + 1):
        for c in range(1, len(HEADERS) + 1):
            x = ws.cell(ri, c); x.font = Font(name="Arial", size=10); x.border = brd
            x.alignment = Alignment(vertical="top", wrap_text=c not in (10, 13))
            if c == 1: x.number_format = "yyyy-mm-dd"
            if c in (7, 8): x.number_format = "0.0000000"
            if c in (10, 13) and x.value:
                x.hyperlink = x.value; x.font = Font(name="Arial", size=10, color="0563C1", underline="single")
    for i, w in enumerate([12, 44, 30, 30, 60, 20, 12, 12, 24, 52, 26, 34, 38], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    n = wb.create_sheet("Catatan")
    n.append(["Item", "Keterangan"])
    n.append(["Diperbarui", dt.datetime.now().strftime("%Y-%m-%d %H:%M")])
    n.append(["Jumlah baris", len(rows)])
    n.append(["Sumber", f"{BASE}/kajian/cari"])
    n.append(["Masjid belum ada di masjid_master.csv", ", ".join(sorted(unknown)) or "-"])
    n.column_dimensions["A"].width = 36; n.column_dimensions["B"].width = 110
    wb.save(OUT / "jadwal_kajian_marimengaji.xlsx")
    with open(OUT / "jadwal_kajian_marimengaji.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(HEADERS)
        for r in rows: w.writerow([r[0].isoformat()] + r[1:])


def write_json(master):
    used = {r[5] for r in RECS}
    m = {}
    for key, v in master.items():
        if v["nama_masjid"] in used:
            m[v["nama_masjid"]] = [v["alamat"], v["kota"], v["latitude"], v["longitude"], v["link_maps"]]
    for n in used - set(m):
        m[n] = ["", "", "", "", ""]
    d = Path("docs"); d.mkdir(exist_ok=True)
    (d / "data.json").write_text(json.dumps({"u": TODAY.isoformat(), "m": m, "r": RECS}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def main():
    cards = scrape_cards()
    if len(cards) < 20:                       # pengaman: jangan timpa file lama dengan hasil kosong
        sys.exit(f"Hanya {len(cards)} kartu terbaca, kemungkinan struktur website berubah. File lama dibiarkan.")
    rows, unknown = build_rows(cards, load_master())
    write_outputs(rows, unknown)
    write_json(load_master())
    print(f"OK: {len(rows)} baris dari {len(cards)} kartu. Masjid baru: {sorted(unknown) or '-'}")


if __name__ == "__main__":
    main()
