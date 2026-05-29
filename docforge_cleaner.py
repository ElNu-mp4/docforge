#!/usr/bin/env python3
"""
docforge_cleaner -- Pembersih Metadata & RSID untuk File .docx
Mengubah dokumen agar mendapat skor risiko copy-paste = 0 dari docforge.

Penggunaan:
  python docforge_cleaner.py input.docx --author "Nama Kamu"
  python docforge_cleaner.py input.docx --author "Nama" --revision 18 --total-time 150
  python docforge_cleaner.py input.docx --rotation-freq 4 --out output.docx
  python docforge_cleaner.py input.docx --dry-run --verbose
"""

import argparse
import random
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from pathlib import Path


# ---------------------------------------------------------------------------
# Namespace map -- hanya untuk core.xml dan app.xml
# ---------------------------------------------------------------------------
NS_CORE = {
    "cp":  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc":  "http://purl.org/dc/elements/1.1/",
    "dct": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
NS_APP = {
    "ep":  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}
for _p, _u in {**NS_CORE, **NS_APP}.items():
    ET.register_namespace(_p, _u)


# ---------------------------------------------------------------------------
# RSID utilities
# ---------------------------------------------------------------------------

def new_rsid():
    return "".join(random.choices("0123456789ABCDEF", k=8))

def generate_rsid_pool(n):
    pool = set()
    while len(pool) < n:
        pool.add(new_rsid())
    return list(pool)


# ---------------------------------------------------------------------------
# Raw-bytes RSID manipulation
# Strategi: TIDAK parse ulang document.xml -- hanya ubah nilai atribut via regex.
# Ini menjaga semua namespace dan struktur XML tetap utuh.
# ---------------------------------------------------------------------------

# Cocokkan paragraf: bisa <w:p> atau <w:p ...atribut...>
_OPEN_PARA_RE = re.compile(rb'<w:p(?:\s[^>]*)?>',  re.DOTALL)
# Cocokkan run: bisa <w:r> atau <w:r ...atribut...>
_OPEN_RUN_RE  = re.compile(rb'<w:r(?:\s[^>]*)?>',  re.DOTALL)
# Noise RSID attributes yang ingin dihapus (bukan rsidR utama)
_NOISE_RE = re.compile(
    rb'\s+w:rsid(?:RPr|Del|Default|RDefault)\s*=\s*"[0-9A-Fa-f]{8}"'
)


def _inject_or_replace_rsid_in_tag(tag_bytes, new_rsid_val):
    """
    Dalam satu tag bytes (mis. b'<w:p w:rsidR="OLD" ...>'),
    ganti atau sisipkan w:rsidR="new_rsid_val".
    Kembalikan tag bytes yang sudah diperbarui.
    """
    r_enc = new_rsid_val.encode()

    # Jika sudah ada w:rsidR, ganti nilainya
    existing = re.search(rb'w:rsidR\s*=\s*"[0-9A-Fa-f]{8}"', tag_bytes)
    if existing:
        return tag_bytes[:existing.start()] + \
               b'w:rsidR="' + r_enc + b'"' + \
               tag_bytes[existing.end():]

    # Belum ada -- sisipkan sebelum penutup '>'
    # Tag bisa diakhiri '>' atau '/>' tapi paragraf/run tidak self-closing
    if tag_bytes.endswith(b'>') and not tag_bytes.endswith(b'/>'):
        inner = tag_bytes[:-1].rstrip()   # hapus '>' dan spasi trailing
        return inner + b' w:rsidR="' + r_enc + b'">'
    return tag_bytes  # tidak dikenal, kembalikan apa adanya


def rotate_rsids_raw(doc_bytes, settings_bytes, rotation_freq):
    """
    1. Temukan semua posisi <w:p...> dan <w:r...>
    2. Tentukan RSID per paragraf berdasarkan rotation_freq
    3. Sinkronkan RSID run dengan paragraf induknya
    4. Rekonstruksi bytes dokumen
    5. Perbarui settings.xml
    """

    # --- Kumpulkan semua match paragraf dan run ---
    para_matches = list(_OPEN_PARA_RE.finditer(doc_bytes))
    run_matches  = list(_OPEN_RUN_RE.finditer(doc_bytes))

    n_paras = len(para_matches)
    if n_paras == 0:
        return doc_bytes, settings_bytes, {"n_paras": 0, "n_sessions": 0,
                                            "n_extra": 0, "rsid_root": "",
                                            "pool_size": 0}

    # --- Buat pool RSID ---
    n_sessions = max(1, (n_paras + rotation_freq - 1) // rotation_freq)
    n_extra    = random.randint(3, 8)
    rsid_pool  = generate_rsid_pool(n_sessions + n_extra)
    rsid_root  = rsid_pool[0]

    # --- Petakan indeks paragraf -> RSID sesi (dengan jitter) ---
    para_to_rsid = {}
    session_i    = 0
    in_session   = 0
    for i in range(n_paras):
        para_to_rsid[i] = rsid_pool[session_i]
        in_session += 1
        jitter    = random.randint(-1, 1)
        effective = max(1, rotation_freq + jitter)
        if in_session >= effective and session_i < n_sessions - 1:
            session_i += 1
            in_session = 0

    # --- Petakan posisi awal paragraf untuk binary search ---
    para_starts = [m.start() for m in para_matches]

    def find_para_for_pos(pos):
        lo, hi, res = 0, len(para_starts) - 1, 0
        while lo <= hi:
            mid = (lo + hi) // 2
            if para_starts[mid] <= pos:
                res = mid
                lo  = mid + 1
            else:
                hi  = mid - 1
        return res

    # --- Kumpulkan semua substitusi (para dan run) ---
    subs = []  # list of (start, end, new_bytes)

    for idx, m in enumerate(para_matches):
        new_r   = para_to_rsid[idx]
        new_tag = _inject_or_replace_rsid_in_tag(m.group(0), new_r)
        # Hapus noise RSID dari tag ini
        new_tag = _NOISE_RE.sub(b"", new_tag)
        subs.append((m.start(), m.end(), new_tag))

    for m in run_matches:
        pidx  = find_para_for_pos(m.start())
        new_r = para_to_rsid.get(pidx, rsid_pool[0])
        new_tag = _inject_or_replace_rsid_in_tag(m.group(0), new_r)
        new_tag = _NOISE_RE.sub(b"", new_tag)
        subs.append((m.start(), m.end(), new_tag))

    # --- Urutkan dan hapus tumpang tindih ---
    subs.sort(key=lambda x: x[0])
    deduped  = []
    last_end = 0
    for start, end, repl in subs:
        if start >= last_end:
            deduped.append((start, end, repl))
            last_end = end

    # --- Rekonstruksi bytes ---
    parts  = []
    cursor = 0
    for start, end, repl in deduped:
        parts.append(doc_bytes[cursor:start])
        parts.append(repl)
        cursor = end
    parts.append(doc_bytes[cursor:])
    new_doc = b"".join(parts)

    # --- Patch settings.xml jika ada ---
    new_settings = settings_bytes
    if settings_bytes:
        used_rsids = sorted(
            set(para_to_rsid.values()) |
            set(rsid_pool[n_sessions : n_sessions + n_extra])
        )

        # Ganti rsidRoot
        new_settings = re.sub(
            rb'(w:rsidRoot\s+w:val\s*=\s*")[0-9A-Fa-f]{8}(")',
            b"\\g<1>" + rsid_root.encode() + b"\\g<2>",
            new_settings,
        )

        # Ganti seluruh blok <w:rsids>
        entries   = b"".join(
            b'\n        <w:rsid w:val="' + r.encode() + b'"/>'
            for r in used_rsids
        )
        root_line = b'\n        <w:rsidRoot w:val="' + rsid_root.encode() + b'"/>'
        new_block = b"<w:rsids>" + root_line + entries + b"\n      </w:rsids>"

        if re.search(rb"<w:rsids>", new_settings):
            new_settings = re.sub(
                rb"<w:rsids>.*?</w:rsids>",
                new_block,
                new_settings,
                flags=re.DOTALL,
            )
        else:
            # Tidak ada blok rsids sama sekali, sisipkan sebelum </w:settings>
            new_settings = new_settings.replace(
                b"</w:settings>",
                b"  " + new_block + b"\n</w:settings>",
            )

    stats = {
        "n_paras":    n_paras,
        "n_sessions": n_sessions,
        "n_extra":    n_extra,
        "rsid_root":  rsid_root,
        "pool_size":  len(rsid_pool),
    }
    return new_doc, new_settings, stats


# ---------------------------------------------------------------------------
# core.xml patching via ET (namespace sederhana, aman)
# ---------------------------------------------------------------------------

def _et_set(root, ns_uri, elem_name, value):
    tag = "{%s}%s" % (ns_uri, elem_name)
    el  = root.find(".//" + tag)
    if el is None:
        el = ET.SubElement(root, tag)
    el.text = value


def _reserialize(root, original_raw):
    body  = ET.tostring(root, encoding="unicode", xml_declaration=False)
    first = original_raw.split(b"\n")[0]
    decl  = (first.decode("utf-8", errors="replace").rstrip() + "\n"
             if first.startswith(b"<?xml")
             else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    return (decl + body).encode("utf-8")


def patch_core_xml(raw, args):
    root   = ET.fromstring(raw)
    author = args.author or "User"

    _et_set(root, NS_CORE["dc"],  "creator",        author)
    _et_set(root, NS_CORE["cp"],  "lastModifiedBy", author)
    _et_set(root, NS_CORE["dc"],  "title",          args.title or "")
    _et_set(root, NS_CORE["dc"],  "subject",        args.subject or "")
    _et_set(root, NS_CORE["cp"],  "keywords",       args.keywords or "")
    _et_set(root, NS_CORE["cp"],  "revision",       str(args.revision))

    now  = datetime.now(timezone.utc)
    days = args.days_spread or random.randint(3, 14)
    created_dt  = now - timedelta(days=days,
                                  hours=random.randint(8, 22),
                                  minutes=random.randint(0, 59))
    modified_dt = now - timedelta(hours=random.randint(0, 8),
                                  minutes=random.randint(0, 59))
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    _et_set(root, NS_CORE["dct"], "created",  created_dt.strftime(fmt))
    _et_set(root, NS_CORE["dct"], "modified", modified_dt.strftime(fmt))
    return _reserialize(root, raw)


def patch_app_xml(raw, args, word_count):
    root = ET.fromstring(raw)
    ep   = NS_APP["ep"]

    def set_tag(name, value):
        tag = "{%s}%s" % (ep, name)
        el  = root.find(".//" + tag)
        if el is None:
            el = ET.SubElement(root, tag)
        el.text = value

    versions = ["16.0000", "16.0100", "16.0200", "16.0300", "15.0000"]
    set_tag("Application", args.app_name or "Microsoft Office Word")
    set_tag("AppVersion",  args.app_version or random.choice(versions))
    set_tag("Company",     args.company or "")
    set_tag("Template",    args.template or "Normal.dotm")

    if args.total_time:
        total_time = args.total_time
    elif word_count > 0:
        mean_wpm  = (args.wpm_min + args.wpm_max) / 2
        std_wpm   = (args.wpm_max - args.wpm_min) / 4
        wpm       = max(args.wpm_min, min(args.wpm_max,
                        random.gauss(mean_wpm, std_wpm)))
        extra_pct = random.uniform(0.25, 0.60)
        total_time = int((word_count / wpm) * (1 + extra_pct))
    else:
        total_time = random.randint(40, 120)

    set_tag("TotalTime", str(max(total_time, 8)))
    return _reserialize(root, raw)


# ---------------------------------------------------------------------------
# Word count dari raw bytes
# ---------------------------------------------------------------------------

def count_words(doc_bytes):
    text = re.sub(rb"<[^>]+>", b" ", doc_bytes)
    return len(text.split())


# ---------------------------------------------------------------------------
# Auto revision
# ---------------------------------------------------------------------------

def auto_revision(wc):
    if wc < 200:  return random.randint(6, 14)
    if wc < 600:  return random.randint(10, 22)
    if wc < 1500: return random.randint(15, 35)
    if wc < 4000: return random.randint(20, 50)
    return random.randint(30, 70)


# ---------------------------------------------------------------------------
# Main cleaner
# ---------------------------------------------------------------------------

def clean_docx(input_path, output_path, args):
    report = {
        "input":      str(input_path),
        "output":     str(output_path),
        "patches":    [],
        "rsid_stats": {},
    }

    with zipfile.ZipFile(input_path, "r") as zf:
        names = zf.namelist()
        files = {n: zf.read(n) for n in names}

    changed    = {}
    doc_raw    = files.get("word/document.xml", b"")
    word_count = count_words(doc_raw)
    report["word_count"] = word_count

    if args.revision is None:
        args.revision = auto_revision(word_count)

    # 1. Rotasi RSID (raw bytes)
    if doc_raw and not args.skip_rsid:
        settings_raw = files.get("word/settings.xml", b"")
        new_doc, new_settings, stats = rotate_rsids_raw(
            doc_raw, settings_raw, args.rotation_freq)
        changed["word/document.xml"] = new_doc
        if new_settings:
            changed["word/settings.xml"] = new_settings
        report["rsid_stats"] = stats
        report["patches"].append(
            "RSID: %d sesi, pool=%d, para=%d" % (
                stats["n_sessions"], stats["pool_size"], stats["n_paras"]))

    # 2. Patch core.xml
    core_raw = files.get("docProps/core.xml")
    if core_raw and not args.skip_core:
        try:
            changed["docProps/core.xml"] = patch_core_xml(core_raw, args)
            report["patches"].append(
                "core.xml: author='%s', revision=%d" % (args.author, args.revision))
        except ET.ParseError as e:
            report["patches"].append("[WARN] core.xml: %s" % e)

    # 3. Patch app.xml
    app_raw = files.get("docProps/app.xml")
    if app_raw and not args.skip_app:
        try:
            changed["docProps/app.xml"] = patch_app_xml(app_raw, args, word_count)
            report["patches"].append(
                "app.xml: TotalTime dari %d kata" % word_count)
        except ET.ParseError as e:
            report["patches"].append("[WARN] app.xml: %s" % e)

    # 4. Tulis ZIP
    if not args.dry_run:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for name in names:
                zf_out.writestr(name, changed.get(name, files[name]))
        report["written"] = True
    else:
        report["written"] = False

    return report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(
        prog="docforge_cleaner",
        description="Bersihkan metadata & RSID .docx agar skor risiko docforge = 0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Contoh:
  python docforge_cleaner.py tugas.docx --author "Budi Santoso"
  python docforge_cleaner.py tugas.docx --author "Ani" --revision 18 --total-time 150
  python docforge_cleaner.py tugas.docx --rotation-freq 3 --out bersih.docx
  python docforge_cleaner.py tugas.docx --dry-run --verbose

Catatan --rotation-freq:
  2-3  : banyak sesi, distribusi RSID sangat merata
  5-7  : default, cocok untuk kebanyakan dokumen
  10+  : sedikit sesi (dokumen pendek, diedit sekali duduk)
        """)

    p.add_argument("file")
    p.add_argument("--out", "-o")

    m = p.add_argument_group("Metadata")
    m.add_argument("--author")
    m.add_argument("--title",       default="")
    m.add_argument("--subject",     default="")
    m.add_argument("--keywords",    default="")
    m.add_argument("--revision",    type=int,   default=None)
    m.add_argument("--total-time",  dest="total_time",  type=int, default=None)
    m.add_argument("--days-spread", dest="days_spread", type=int, default=None)
    m.add_argument("--app-name",    dest="app_name",    default=None)
    m.add_argument("--app-version", dest="app_version", default=None)
    m.add_argument("--company",     default="")
    m.add_argument("--template",    default=None)

    r = p.add_argument_group("RSID")
    r.add_argument("--rotation-freq", dest="rotation_freq", type=int, default=5, metavar="N")
    r.add_argument("--seed",          type=int, default=None)

    w = p.add_argument_group("Kecepatan mengetik")
    w.add_argument("--wpm-min", dest="wpm_min", type=float, default=15.0)
    w.add_argument("--wpm-max", dest="wpm_max", type=float, default=35.0)

    s = p.add_argument_group("Skip")
    s.add_argument("--skip-rsid", action="store_true")
    s.add_argument("--skip-core", action="store_true")
    s.add_argument("--skip-app",  action="store_true")

    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--verbose", "-v", action="store_true")
    return p


def main():
    parser = build_parser()
    args   = parser.parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    input_path = Path(args.file)
    if not input_path.exists():
        print("[ERROR] File tidak ditemukan: %s" % input_path, file=sys.stderr)
        sys.exit(1)
    if input_path.suffix.lower() not in (".docx", ".dotx"):
        print("[ERROR] Hanya mendukung .docx/.dotx", file=sys.stderr)
        sys.exit(1)

    output_path = (Path(args.out) if args.out
                   else input_path.with_name(input_path.stem + "_clean.docx"))

    print()
    print("  Input   : %s" % input_path)
    if args.dry_run:
        print("  Mode    : DRY-RUN")
    else:
        print("  Output  : %s" % output_path)
    print("  Author  : %s" % (args.author or "(default)"))
    print("  Rotasi  : setiap %d paragraf" % args.rotation_freq)
    print()

    report = clean_docx(input_path, output_path, args)

    print("  Perubahan:")
    for patch in report["patches"]:
        print("    - %s" % patch)

    if args.verbose and report.get("rsid_stats"):
        s = report["rsid_stats"]
        print()
        print("  Detail RSID:")
        print("    Paragraf    : %s" % s.get("n_paras",    "-"))
        print("    Sesi        : %s" % s.get("n_sessions", "-"))
        print("    Pool total  : %s" % s.get("pool_size",  "-"))
        print("    rsidRoot    : %s" % s.get("rsid_root",  "-"))
        print("    Kata (est.) : %s" % report.get("word_count", "-"))
        print("    Revisi      : %s" % args.revision)

    if not args.dry_run:
        size_in  = input_path.stat().st_size
        size_out = output_path.stat().st_size
        print()
        print("  Ukuran  : %s B -> %s B" % (f"{size_in:,}", f"{size_out:,}"))
        print("  Selesai : %s" % output_path)
    else:
        print()
        print("  (Dry-run, tidak ada file ditulis)")
    print()


if __name__ == "__main__":
    main()
