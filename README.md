# docforge

> Analisis metadata & indikator copy-paste pada file Microsoft Office

`docforge` adalah CLI toolkit berbasis Python yang membedah metadata internal file Office untuk mendeteksi pola copy-paste, menelusuri riwayat pengeditan, dan menghasilkan skor risiko yang dapat dibaca manusia maupun mesin. Dilengkapi dengan cleaner untuk membersihkan metadata dan merotasi RSID.

---

## Daftar Isi

- [Latar Belakang](#latar-belakang)
- [Format yang Didukung](#format-yang-didukung)
- [Instalasi](#instalasi)
- [Penggunaan](#penggunaan)
  - [docforge.py — Analyzer](#docforgepy--analyzer)
  - [docforge_cleaner.py — Cleaner](#docforge_cleanerpy--cleaner)
- [Cara Kerja](#cara-kerja)
- [Output](#output)
- [Interpretasi Hasil](#interpretasi-hasil)
- [Keterbatasan](#keterbatasan)

---

## Latar Belakang

File Office modern (`.docx`, `.pptx`, `.xlsx`) sebenarnya adalah arsip ZIP berisi kumpulan XML. Di dalam XML tersebut tersimpan metadata yang jarang diketahui pengguna biasa, termasuk:

- **Siapa yang membuat** dan **siapa yang terakhir mengedit** dokumen
- **Berapa kali** dokumen disimpan
- **RSID** (*Revision Save ID*) — ID unik yang Word assign ke setiap sesi pengetikan

Dengan menganalisis metadata ini, `docforge` dapat mengidentifikasi pola yang mengindikasikan bahwa konten dokumen tidak ditulis secara organik, melainkan dipindahkan dari sumber lain dalam satu operasi besar.

---

## Format yang Didukung

| Ekstensi | Format | Analisis RSID |
|---|---|---|
| `.docx` `.dotx` | Word Document | ✅ Penuh |
| `.pptx` `.potx` | PowerPoint | ✅ Metadata + slide count |
| `.xlsx` `.xlsm` | Excel | ✅ Metadata + sheet info |

---

## Instalasi

**Prasyarat:** Python 3.8 atau lebih baru.

```bash
# 1. Install dependency
pip install rich

# 2. Clone repo
git clone https://github.com/ElNu-mp4/docforge.git
cd docforge

# 3. (Opsional) Jadikan executable global di Linux/macOS
chmod +x docforge.py
sudo cp docforge.py /usr/local/bin/docforge
```

Tidak ada dependency lain — semua modul seperti `zipfile`, `xml.etree`, dan `json` sudah bawaan Python.

---

## Penggunaan

### docforge.py — Analyzer

```bash
python3 docforge.py <file> [opsi]
```

| Argumen/Opsi | Deskripsi |
|---|---|
| `file` | Path ke file Office yang ingin dianalisis |
| `--json` | Tambahkan output JSON di bawah dashboard |
| `--out <path>` | Simpan laporan JSON ke file |
| `--no-dashboard` | Hanya tampilkan JSON, tanpa dashboard terminal |
| `-h`, `--help` | Tampilkan bantuan |

**Contoh:**

```bash
# Analisis dasar
python3 docforge.py laporan.docx

# Simpan laporan JSON
python3 docforge.py laporan.docx --out hasil_analisis.json

# JSON only (cocok untuk pipeline/scripting)
python3 docforge.py laporan.docx --no-dashboard

# Analisis file PowerPoint / Excel
python3 docforge.py presentasi.pptx
python3 docforge.py data.xlsx --out laporan.json
```

**Contoh output terminal:**

```
──────── docforge  •  laporan.docx ────────

╭─────────────── Skor Risiko Copy-Paste ───────────────╮
│                                                       │
│  SEDANG   ██████████░░░░░░░░░░   50/100              │
│                                                       │
│  ⚠  Dokumen dibuat oleh 'Sandy', diedit oleh 'M S I' │
│  ⚠  Hanya 3x disimpan (sangat sedikit)               │
│  ⚠  RSID dominan 95.7%: hampir semua teks satu sesi  │
│                                                       │
╰───────────────────────────────────────────────────────╯

📄  Metadata Dokumen
┌────────────────────────┬──────────────────────────┐
│ Dibuat oleh            │ Sandy Kurniawan           │
│ Terakhir diedit        │ M S I                     │
│ Revisi                 │ 3                         │
│ Tanggal dibuat         │ 27 Feb 2026, 05:41 UTC    │
│ Terakhir disimpan      │ 02 May 2026, 23:17 UTC    │
│ Aplikasi               │ Microsoft Office Word     │
└────────────────────────┴──────────────────────────┘

📈  Distribusi RSID (level run)
  00C81389    1080   95.7%   ███████████████████░
  00B96628      11    1.0%   ░░░░░░░░░░░░░░░░░░░░
  003030D0       7    0.6%   ░░░░░░░░░░░░░░░░░░░░
```

---

### docforge_cleaner.py — Cleaner

Membersihkan metadata dan merotasi RSID pada file `.docx` agar skor risiko docforge = 0.

```bash
python3 docforge_cleaner.py <file> [opsi]
```

| Opsi | Deskripsi |
|---|---|
| `--author` | Nama author yang akan ditanamkan ke metadata |
| `--revision` | Jumlah revisi (default: dihitung otomatis dari word count) |
| `--total-time` | Total waktu edit dalam menit (default: dihitung dari word count) |
| `--rotation-freq N` | Frekuensi rotasi RSID per N paragraf (default: 5) |
| `--out`, `-o` | Path output (default: `input_clean.docx`) |
| `--dry-run` | Simulasi tanpa menulis file |
| `--verbose`, `-v` | Tampilkan detail RSID |
| `--skip-rsid` | Lewati rotasi RSID |
| `--skip-core` | Lewati patch core.xml |
| `--skip-app` | Lewati patch app.xml |

**Contoh:**

```bash
# Bersihkan dengan author baru
python3 docforge_cleaner.py tugas.docx --author "Nama Kamu"

# Tentukan revision dan total waktu edit secara manual
python3 docforge_cleaner.py tugas.docx --author "Nama" --revision 18 --total-time 150

# Atur frekuensi rotasi RSID
python3 docforge_cleaner.py tugas.docx --rotation-freq 4 --out bersih.docx

# Dry-run
python3 docforge_cleaner.py tugas.docx --dry-run --verbose
```

**Opsi `--rotation-freq`:**

| Nilai | Efek |
|---|---|
| 2–3 | Banyak sesi, distribusi RSID sangat merata |
| 5–7 | Default, cocok untuk kebanyakan dokumen |
| 10+ | Sedikit sesi (dokumen pendek) |

---

## Cara Kerja

### Analyzer (`docforge.py`)

File Office adalah arsip ZIP. `docforge` membukanya tanpa mengekstrak ke disk, lalu membaca tiga lapisan metadata:

**1. Core Metadata (`docProps/core.xml`)**

| Field XML | Arti |
|---|---|
| `dc:creator` | Username yang membuat dokumen pertama kali |
| `cp:lastModifiedBy` | Username yang terakhir menyimpan |
| `cp:revision` | Jumlah akumulasi operasi simpan |
| `dcterms:created` | Timestamp pembuatan |
| `dcterms:modified` | Timestamp penyimpanan terakhir |

**2. App Metadata (`docProps/app.xml`)**

| Field XML | Arti |
|---|---|
| `TotalTime` | Total menit dokumen pernah dibuka dalam mode edit |
| `Words` | Jumlah kata |
| `Pages` | Jumlah halaman |
| `Application` | Aplikasi yang dipakai |
| `Template` | Template yang digunakan saat dibuat |

**3. RSID Analysis — khusus `.docx`**

RSID (*Revision Save ID*) adalah angka heksadesimal 8 karakter yang Word assign setiap sesi pengeditan baru, dicatat pada `w:rsidR` di elemen `<w:p>` (paragraf) dan `<w:r>` (run/teks).

| Indikator | Metode Deteksi | Interpretasi |
|---|---|---|
| **RSID asing** | RSID di `document.xml` tidak ada di `settings.xml` | Konten dipaste dari dokumen Word lain |
| **Dominansi RSID** | % RSID terbanyak dari total run | >80% = sebagian besar teks dari satu sesi |
| **Mismatch paragraf↔run** | Paragraf RSID ≠ RSID run di dalamnya | Struktur dari satu sesi, isi dari sesi berbeda |
| **rsidRoot** | Nilai `<w:rsidRoot>` | Identitas asal dokumen / sesi pertama |

**4. Scoring Risiko**

| Kondisi | Tambahan Skor |
|---|---|
| Creator ≠ lastModifiedBy | +15 |
| Revisi = 1 | +20 |
| Revisi ≤ 3 | +10 |
| RSID dominan ≥ 95% | +25 |
| RSID dominan ≥ 80% | +15 |
| Ada RSID asing | +30 |
| Mismatch paragraf↔run ≥ 20% | +15 |
| Kecepatan edit > 100 kata/menit | +20 |
| Kecepatan edit > 50 kata/menit | +10 |

| Skor | Level | Warna |
|---|---|---|
| 0 – 39 | RENDAH | Hijau |
| 40 – 69 | SEDANG | Kuning |
| 70 – 100 | TINGGI | Merah |

### Cleaner (`docforge_cleaner.py`)

- Merotasi RSID di `document.xml` dan `settings.xml` via raw bytes
- Menyinkronkan RSID run dengan paragraf induknya
- Menambahkan jitter acak antar sesi untuk distribusi yang natural
- Patch `core.xml`: author, timestamps, revision
- Patch `app.xml`: aplikasi, total waktu edit (dihitung dari word count)

---

## Output

### Dashboard Terminal

1. **Banner risiko** — skor, level, dan daftar flag
2. **Tabel metadata dokumen** — informasi identitas dan aplikasi
3. **Tabel statistik** — kata, halaman, waktu, kecepatan
4. **Tabel RSID** — ringkasan analisis sesi (khusus `.docx`)
5. **Distribusi RSID** — bar chart ASCII per sesi
6. **Contoh mismatch** — cuplikan teks dengan RSID tidak konsisten

### Laporan JSON

```json
{
  "file": "laporan.docx",
  "format": "docx",
  "analyzed": "2026-05-03T06:26:33+00:00",
  "risk": {
    "score": 50,
    "level": "SEDANG",
    "flags": ["Dokumen dibuat oleh 'Sandy', diedit oleh 'M S I'", "..."],
    "notes": []
  },
  "core_metadata": {
    "creator": "Sandy Kurniawan",
    "last_modified_by": "M S I",
    "revision": "3",
    "created": "2026-02-27T05:41:00Z",
    "modified": "2026-05-02T23:17:00Z"
  },
  "app_metadata": {
    "application": "Microsoft Office Word",
    "total_time": "426",
    "words": "8083",
    "pages": "50"
  },
  "rsid_analysis": {
    "rsid_root": "0092077F",
    "registered_count": 107,
    "unique_run_rsids": 12,
    "dominant_rsid": "00C81389",
    "dominant_count": 1080,
    "dominant_pct": 95.7,
    "foreign_rsids": [],
    "total_para": 516,
    "mismatch_para": 5,
    "mismatch_pct": 1.0
  }
}
```

---

## Interpretasi Hasil

### RSID Dominan Tinggi (>80%)

Bukan otomatis berarti plagiarisme. Ada dua skenario umum:

**Skenario wajar:** Mahasiswa mengisi template dosen dalam satu sesi kerja panjang. Semua teks yang diketik dalam sesi itu mendapat RSID yang sama. Ini normal dan sering terjadi.

**Skenario mencurigakan:** Seseorang mem-paste konten dari luar Word (Google Docs, Notepad, web) lalu menyimpannya. Semua teks yang dipaste mendapat RSID sesi saat itu, menyebabkan dominansi ekstrem.

Bedakan keduanya dengan melihat **kecepatan edit** (kata/menit) dan **jumlah revisi**. Jika dominansi tinggi + revisi sangat sedikit + kecepatan tidak wajar → lebih mencurigakan.

### RSID Asing Ditemukan

Ini adalah **indikator terkuat**. RSID asing muncul ketika seseorang membuka dua dokumen Word berbeda di komputer yang sama, lalu meng-copy-paste antar keduanya. Teks yang dipaste membawa RSID dari dokumen asal yang tidak terdaftar di `settings.xml` dokumen tujuan.

### Creator ≠ Last Modified By

Sangat umum dan tidak selalu bermasalah — misalnya dosen membagikan template, lalu mahasiswa mengisinya. Tapi dalam konteks pengumpulan tugas antar-mahasiswa, flag ini perlu dicermati lebih lanjut.

### Mismatch Paragraf ↔ Run

Terjadi ketika struktur paragraf (tanda `¶`) berasal dari satu sesi, tetapi isi teks di dalamnya ditulis di sesi berbeda. Ini bisa terjadi karena penghapusan dan pengetikan ulang, atau karena paste yang menimpa teks lama.

---

## Keterbatasan

- **Bukan bukti hukum.** Hasil analisis bersifat indikatif, bukan konklusif. Metadata bisa dimanipulasi secara manual.
- **RSID hanya ada di `.docx`.** Format `.pptx` dan `.xlsx` tidak memiliki sistem RSID setara.
- **Tidak mendeteksi parafrase.** `docforge` menganalisis metadata teknis, bukan kesamaan konten. Untuk plagiarisme berbasis teks, gunakan tools seperti Turnitin.
- **Paste dari luar Word tidak meninggalkan RSID asing.** Jika konten disalin dari browser, PDF, atau Google Docs, tidak ada RSID asing — hanya dominansi RSID satu sesi.
- **File yang disimpan ulang Word Online** dapat menormalisasi RSID, menghilangkan jejak sesi sebelumnya.

---

## Project Structure

```
docforge/
├── README.md
├── requirements.txt
├── docforge.py           # analyzer
└── docforge_cleaner.py   # cleaner
```

## Author

Elang N