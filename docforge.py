#!/usr/bin/env python3
"""
docforge — Office File Metadata & Copy-Paste Analyzer
Supports: .docx, .pptx, .xlsx, .odt, .ods, .odp
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
import argparse
import json
import re
from pathlib import Path
from collections import Counter
from datetime import datetime, timezone

# ── Rich UI ───────────────────────────────────────────────────────────────────
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.columns import Columns
from rich.text import Text
from rich.rule import Rule
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich import box
from rich.padding import Padding

console = Console()

# ── Namespaces ─────────────────────────────────────────────────────────────────
NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_CP  = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC  = "http://purl.org/dc/elements/1.1/"
NS_DCT = "http://purl.org/dc/terms/"
NS_EP  = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_PR  = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"


def tag(ns, name):
    return f"{{{ns}}}{name}"


# ── XML helpers ────────────────────────────────────────────────────────────────

def read_xml_from_zip(zf: zipfile.ZipFile, path: str):
    """Read and parse an XML file inside the ZIP, returns None if not found."""
    try:
        with zf.open(path) as f:
            return ET.parse(f).getroot()
    except (KeyError, ET.ParseError):
        return None


def text_of(root, *tag_path):
    """Walk a tag path, return text or ''."""
    el = root
    for t in tag_path:
        el = el.find(t)
        if el is None:
            return ""
    return (el.text or "").strip()


# ── Metadata extraction ────────────────────────────────────────────────────────

def parse_core(zf: zipfile.ZipFile) -> dict:
    root = read_xml_from_zip(zf, "docProps/core.xml")
    if root is None:
        return {}

    def g(ns, name):
        el = root.find(f"{{{ns}}}{name}")
        return (el.text or "").strip() if el is not None else ""

    created_raw  = g(NS_DCT, "created")
    modified_raw = g(NS_DCT, "modified")

    return {
        "title":            g(NS_DC,  "title"),
        "subject":          g(NS_DC,  "subject"),
        "creator":          g(NS_DC,  "creator"),
        "last_modified_by": g(NS_CP,  "lastModifiedBy"),
        "revision":         g(NS_CP,  "revision"),
        "keywords":         g(NS_CP,  "keywords"),
        "description":      g(NS_DC,  "description"),
        "created":          created_raw,
        "modified":         modified_raw,
    }


def parse_app(zf: zipfile.ZipFile) -> dict:
    root = read_xml_from_zip(zf, "docProps/app.xml")
    if root is None:
        return {}

    def g(name):
        el = root.find(f"{{{NS_EP}}}{name}")
        return (el.text or "").strip() if el is not None else ""

    return {
        "application":  g("Application"),
        "app_version":  g("AppVersion"),
        "total_time":   g("TotalTime"),
        "pages":        g("Pages"),
        "words":        g("Words"),
        "characters":   g("Characters"),
        "paragraphs":   g("Paragraphs"),
        "slides":       g("Slides"),
        "notes":        g("Notes"),
        "company":      g("Company"),
        "template":     g("Template"),
    }


# ── DOCX RSID analysis ────────────────────────────────────────────────────────

def analyze_docx_rsids(zf: zipfile.ZipFile) -> dict:
    """Full RSID analysis for DOCX files."""

    # 1. Registered RSIDs from settings.xml
    settings = read_xml_from_zip(zf, "word/settings.xml")
    registered = set()
    rsid_root  = None
    if settings is not None:
        for el in settings.iter(tag(NS_W, "rsid")):
            v = el.get(tag(NS_W, "val"))
            if v:
                registered.add(v)
        root_el = settings.find(f".//{tag(NS_W, 'rsidRoot')}")
        if root_el is not None:
            rsid_root = root_el.get(tag(NS_W, "val"))
            registered.add(rsid_root)

    # 2. RSIDs used in document.xml
    doc = read_xml_from_zip(zf, "word/document.xml")
    if doc is None:
        return {"error": "document.xml not found"}

    para_rsids = Counter()   # paragraph-level
    run_rsids  = Counter()   # run-level (text content)
    foreign    = set()       # in document but not in settings

    # ── paragraph level ──
    for para in doc.iter(tag(NS_W, "p")):
        r = para.get(tag(NS_W, "rsidR"))
        if r:
            para_rsids[r] += 1
            if registered and r not in registered:
                foreign.add(r)

    # ── run level ──
    for run in doc.iter(tag(NS_W, "r")):
        r = run.get(tag(NS_W, "rsidR"))
        if r:
            run_rsids[r] += 1
            if registered and r not in registered:
                foreign.add(r)

    # 3. Para ↔ Run RSID mismatch (paste indicator)
    total_para    = 0
    mismatch_para = 0
    mismatch_examples = []

    for para in doc.iter(tag(NS_W, "p")):
        p_rsid = para.get(tag(NS_W, "rsidR"))
        if not p_rsid:
            continue
        text = "".join(t.text or "" for t in para.iter(tag(NS_W, "t"))).strip()
        if not text:
            continue
        total_para += 1

        run_rsids_in_para = {
            run.get(tag(NS_W, "rsidR"))
            for run in para.findall(tag(NS_W, "r"))
            if run.get(tag(NS_W, "rsidR"))
        }
        if run_rsids_in_para and p_rsid not in run_rsids_in_para:
            mismatch_para += 1
            if len(mismatch_examples) < 3:
                mismatch_examples.append({
                    "para_rsid": p_rsid,
                    "run_rsid":  list(run_rsids_in_para)[0],
                    "text":      text[:60],
                })

    # 4. Dominant run RSID
    total_runs = sum(run_rsids.values())
    dominant_rsid, dominant_count = run_rsids.most_common(1)[0] if run_rsids else (None, 0)
    dominant_pct = (dominant_count / total_runs * 100) if total_runs else 0

    return {
        "rsid_root":        rsid_root,
        "registered_count": len(registered),
        "para_rsids":       dict(para_rsids.most_common(10)),
        "run_rsids":        dict(run_rsids.most_common(10)),
        "run_rsids_full":   run_rsids,
        "total_runs":       total_runs,
        "dominant_rsid":    dominant_rsid,
        "dominant_count":   dominant_count,
        "dominant_pct":     dominant_pct,
        "foreign_rsids":    list(foreign),
        "total_para":       total_para,
        "mismatch_para":    mismatch_para,
        "mismatch_pct":     (mismatch_para / total_para * 100) if total_para else 0,
        "mismatch_examples":mismatch_examples,
        "unique_run_rsids": len(run_rsids),
    }


# ── PPTX / XLSX metadata hints ────────────────────────────────────────────────

def analyze_pptx(zf: zipfile.ZipFile) -> dict:
    """Basic PPTX-specific metadata."""
    slides = [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml", n)]
    masters = [n for n in zf.namelist() if "slideMaster" in n]

    # Count unique authors in revision list if present
    rev_root = read_xml_from_zip(zf, "ppt/revisionInfo.xml")
    return {
        "slide_count":  len(slides),
        "master_count": len(masters),
        "has_revision_info": rev_root is not None,
    }


def analyze_xlsx(zf: zipfile.ZipFile) -> dict:
    """Basic XLSX-specific metadata."""
    sheets = [n for n in zf.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n)]
    wb = read_xml_from_zip(zf, "xl/workbook.xml")
    sheet_names = []
    if wb is not None:
        for sh in wb.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
            name = sh.get("name")
            if name:
                sheet_names.append(name)
    return {
        "sheet_count": len(sheets),
        "sheet_names": sheet_names,
    }


# ── Risk scoring ──────────────────────────────────────────────────────────────

def compute_risk(core: dict, app: dict, rsid: dict, fmt: str) -> dict:
    score   = 0
    flags   = []
    notes   = []

    creator  = core.get("creator", "")
    last_mod = core.get("last_modified_by", "")
    revision = int(core.get("revision") or 0)
    total_time = int(app.get("total_time") or 0)

    # ── Creator ≠ last modifier ──
    if creator and last_mod and creator.strip().lower() != last_mod.strip().lower():
        score += 15
        flags.append(f"Dokumen dibuat oleh '{creator}', diedit oleh '{last_mod}'")

    # ── Very few revisions ──
    if revision == 1:
        score += 20
        flags.append(f"Hanya 1x disimpan — kemungkinan langsung di-paste")
    elif revision <= 3:
        score += 10
        flags.append(f"Hanya {revision}x disimpan (sangat sedikit)")

    # ── DOCX-specific RSID checks ──
    if fmt == "docx" and rsid and not rsid.get("error"):
        dom_pct = rsid.get("dominant_pct", 0)
        foreign = rsid.get("foreign_rsids", [])
        mis_pct = rsid.get("mismatch_pct", 0)

        # Dominant RSID
        if dom_pct >= 95:
            score += 25
            flags.append(f"RSID dominan {dom_pct:.1f}%: hampir semua teks dari satu sesi")
        elif dom_pct >= 80:
            score += 15
            flags.append(f"RSID dominan {dom_pct:.1f}%: sebagian besar teks dari satu sesi")
        elif dom_pct >= 60:
            score += 8
            notes.append(f"RSID dominan {dom_pct:.1f}%: sebagian teks mungkin di-paste")

        # Foreign RSIDs (strongest indicator of cross-doc paste)
        if foreign:
            score += 30
            flags.append(f"{len(foreign)} RSID asing ditemukan — konten dari dokumen lain!")

        # Para-run mismatch
        if mis_pct >= 20:
            score += 15
            flags.append(f"Mismatch RSID paragraf↔run {mis_pct:.1f}%: indikator paste")
        elif mis_pct >= 5:
            score += 5
            notes.append(f"Mismatch RSID paragraf↔run {mis_pct:.1f}%")

    # ── Edit time vs word count ──
    words = int(app.get("words") or 0)
    if words > 0 and total_time > 0:
        wpm = words / total_time
        if wpm > 100:
            score += 20
            flags.append(f"Kecepatan edit {wpm:.0f} kata/menit — tidak wajar untuk pengetikan manual")
        elif wpm > 50:
            score += 10
            notes.append(f"Kecepatan edit {wpm:.0f} kata/menit — cukup tinggi")

    # ── Cap & level ──
    score = min(score, 100)
    if score >= 70:
        level = "TINGGI"
        color = "red"
    elif score >= 40:
        level = "SEDANG"
        color = "yellow"
    else:
        level = "RENDAH"
        color = "green"

    return {"score": score, "level": level, "color": color, "flags": flags, "notes": notes}


# ── Terminal dashboard ────────────────────────────────────────────────────────

def fmt_dt(raw: str) -> str:
    if not raw:
        return "—"
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        return dt.strftime("%d %b %Y, %H:%M UTC")
    except Exception:
        return raw


def render_dashboard(path: Path, core: dict, app: dict, rsid: dict, risk: dict, fmt: str, extra: dict):
    console.print()
    console.rule(f"[bold cyan]docforge[/bold cyan]  [dim]•  {path.name}[/dim]", style="cyan")
    console.print()

    # ── Risk banner ──
    color = risk["color"]
    score = risk["score"]
    bar   = "█" * (score // 5) + "░" * (20 - score // 5)
    level = risk["level"]

    console.print(Panel(
        f"[bold {color}]{level}[/bold {color}]   [{color}]{bar}[/{color}]   [bold]{score}/100[/bold]\n\n"
        + ("\n".join(f"  [red]⚠[/red]  {f}" for f in risk["flags"]) or "  [green]Tidak ada indikator mencurigakan[/green]")
        + ("\n\n" + "\n".join(f"  [yellow]ℹ[/yellow]  {n}" for n in risk["notes"]) if risk["notes"] else ""),
        title="[bold]Skor Risiko Copy-Paste[/bold]",
        border_style=color,
        padding=(1, 2),
    ))
    console.print()

    # ── Core metadata table ──
    t = Table(title="Metadata Dokumen", box=box.ROUNDED, border_style="cyan", show_lines=True)
    t.add_column("Field", style="bold", width=22)
    t.add_column("Nilai")

    def row(label, val, hi=False):
        style = "bold yellow" if hi else ""
        t.add_row(label, Text(val or "—", style=style))

    row("Format",          fmt.upper())
    row("Dibuat oleh",     core.get("creator",""),      hi=bool(core.get("creator")))
    row("Terakhir diedit", core.get("last_modified_by",""), hi=bool(core.get("last_modified_by")))
    row("Revisi",          core.get("revision",""),     hi=(int(core.get("revision") or 0) <= 3))
    row("Tanggal dibuat",  fmt_dt(core.get("created","")))
    row("Terakhir disimpan", fmt_dt(core.get("modified","")))
    row("Judul",           core.get("title",""))
    row("Aplikasi",        f"{app.get('application','')} v{app.get('app_version','')}")
    row("Template",        app.get("template",""))
    row("Perusahaan",      app.get("company",""))
    console.print(t)
    console.print()

    # ── Stats table ──
    t2 = Table(title="Statistik Dokumen", box=box.ROUNDED, border_style="blue", show_lines=True)
    t2.add_column("Metrik", style="bold", width=22)
    t2.add_column("Nilai")

    total_time = int(app.get("total_time") or 0)
    words      = int(app.get("words") or 0)

    t2.add_row("Total waktu edit",   f"{total_time} menit ({total_time//60}j {total_time%60}m)" if total_time else "—")
    t2.add_row("Jumlah kata",        app.get("words","—"))
    t2.add_row("Jumlah karakter",    app.get("characters","—"))

    if fmt == "docx":
        t2.add_row("Halaman",        app.get("pages","—"))
        t2.add_row("Paragraf",       app.get("paragraphs","—"))
    elif fmt == "pptx":
        t2.add_row("Slide",          str(extra.get("slide_count","—")))
    elif fmt == "xlsx":
        t2.add_row("Sheet",          str(extra.get("sheet_count","—")))
        names = ", ".join(extra.get("sheet_names",[]))
        if names:
            t2.add_row("Nama Sheet",  names)

    if words and total_time:
        t2.add_row("Kecepatan edit", f"{words/total_time:.1f} kata/menit")

    console.print(t2)

    # ── RSID analysis (DOCX only) ──
    if fmt == "docx" and rsid and not rsid.get("error"):
        console.print()
        t3 = Table(title="Analisis RSID (Word Save Sessions)", box=box.ROUNDED, border_style="magenta", show_lines=True)
        t3.add_column("Metrik", style="bold", width=30)
        t3.add_column("Nilai")

        dom_pct  = rsid["dominant_pct"]
        dom_col  = "red" if dom_pct >= 80 else ("yellow" if dom_pct >= 60 else "green")
        for_rsid = rsid["foreign_rsids"]

        t3.add_row("rsidRoot (asal template)",  rsid.get("rsid_root") or "—")
        t3.add_row("RSID terdaftar",            str(rsid["registered_count"]))
        t3.add_row("RSID unik (level run)",     str(rsid["unique_run_rsids"]))
        t3.add_row("RSID dominan",              Text(f"{rsid['dominant_rsid']}  ({rsid['dominant_count']} runs)", style=dom_col))
        t3.add_row("Dominansi RSID (%)",        Text(f"{dom_pct:.1f}%", style=dom_col))
        t3.add_row("RSID asing (cross-doc)",    Text(str(len(for_rsid)) + (" ⚠" if for_rsid else " ✓"), style="red" if for_rsid else "green"))
        t3.add_row("Total paragraf berisi teks",str(rsid["total_para"]))
        t3.add_row("Mismatch para↔run RSID",    f"{rsid['mismatch_para']} ({rsid['mismatch_pct']:.1f}%)")
        console.print(t3)

        # Top RSIDs breakdown
        run_counter = rsid["run_rsids_full"]
        total_runs  = rsid["total_runs"]
        if run_counter:
            console.print()
            t4 = Table(title="Distribusi RSID (level run)", box=box.SIMPLE_HEAD, border_style="dim")
            t4.add_column("RSID",  width=12)
            t4.add_column("Runs",  justify="right", width=8)
            t4.add_column("Persen", justify="right", width=8)
            t4.add_column("Bar", width=30)

            for rsid_val, cnt in run_counter.most_common(8):
                pct     = cnt / total_runs * 100
                bar_len = int(pct / 5)
                bar_str = "█" * bar_len + "░" * (20 - bar_len)
                col     = "red" if pct >= 80 else ("yellow" if pct >= 40 else "cyan")
                foreign_mark = " [red]![/red]" if rsid_val in for_rsid else ""
                t4.add_row(
                    f"[bold]{rsid_val}[/bold]{foreign_mark}",
                    str(cnt),
                    f"{pct:.1f}%",
                    Text(bar_str, style=col),
                )
            console.print(t4)

        # Mismatch examples
        if rsid.get("mismatch_examples"):
            console.print()
            console.print("[bold]Contoh mismatch paragraf ↔ run:[/bold]")
            for ex in rsid["mismatch_examples"]:
                console.print(f"  [dim]Para RSID:[/dim] {ex['para_rsid']}  [dim]Run RSID:[/dim] {ex['run_rsid']}")
                console.print(f"  [italic dim]  → \"{ex['text']}\"[/italic dim]")

        # Foreign RSID list
        if for_rsid:
            console.print()
            console.print(f"[bold red]RSID Asing ditemukan ({len(for_rsid)}):[/bold red]")
            for r in for_rsid[:10]:
                console.print(f"  [red]• {r}[/red]")

    console.print()
    console.rule(style="dim")
    console.print()


# ── JSON report ───────────────────────────────────────────────────────────────

def build_json_report(path: Path, core: dict, app: dict, rsid: dict, risk: dict, fmt: str, extra: dict) -> dict:
    report = {
        "file":     str(path),
        "format":   fmt,
        "analyzed": datetime.now(timezone.utc).isoformat(),
        "risk":     {
            "score": risk["score"],
            "level": risk["level"],
            "flags": risk["flags"],
            "notes": risk["notes"],
        },
        "core_metadata": core,
        "app_metadata":  app,
    }
    if fmt == "docx" and rsid:
        r = dict(rsid)
        r.pop("run_rsids_full", None)  # not JSON-serializable as Counter
        report["rsid_analysis"] = r
    if extra:
        report["format_metadata"] = extra
    return report


# ── Main ─────────────────────────────────────────────────────────────────────

SUPPORTED = {
    ".docx": "docx",
    ".dotx": "docx",
    ".pptx": "pptx",
    ".potx": "pptx",
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
}


def detect_format(path: Path) -> str:
    return SUPPORTED.get(path.suffix.lower(), "unknown")


def analyze(path: Path) -> dict:
    fmt = detect_format(path)
    if fmt == "unknown":
        console.print(f"[red]Format tidak didukung: {path.suffix}[/red]")
        console.print(f"Format yang didukung: {', '.join(SUPPORTED.keys())}")
        sys.exit(1)

    if not path.exists():
        console.print(f"[red]File tidak ditemukan: {path}[/red]")
        sys.exit(1)

    with zipfile.ZipFile(path, "r") as zf:
        core  = parse_core(zf)
        app   = parse_app(zf)
        rsid  = analyze_docx_rsids(zf) if fmt == "docx" else {}
        extra = {}
        if fmt == "pptx":
            extra = analyze_pptx(zf)
        elif fmt == "xlsx":
            extra = analyze_xlsx(zf)

    risk = compute_risk(core, app, rsid, fmt)
    return {"core": core, "app": app, "rsid": rsid, "risk": risk, "fmt": fmt, "extra": extra}


def main():
    parser = argparse.ArgumentParser(
        prog="docforge",
        description="Analisis metadata & indikator copy-paste pada file Office",
    )
    parser.add_argument("file", help="Path ke file Office (.docx/.pptx/.xlsx dll.)")
    parser.add_argument("--json",    action="store_true", help="Output sebagai JSON")
    parser.add_argument("--out",     help="Simpan laporan JSON ke file ini")
    parser.add_argument("--no-dashboard", action="store_true", help="Sembunyikan dashboard, hanya tampilkan JSON")
    args = parser.parse_args()

    path = Path(args.file)

    with Progress(SpinnerColumn(), TextColumn("[cyan]Menganalisis {task.description}…"), transient=True) as p:
        t = p.add_task(path.name)
        data = analyze(path)
        p.update(t, completed=True)

    core  = data["core"]
    app   = data["app"]
    rsid  = data["rsid"]
    risk  = data["risk"]
    fmt   = data["fmt"]
    extra = data["extra"]

    if not args.no_dashboard:
        render_dashboard(path, core, app, rsid, risk, fmt, extra)

    if args.json or args.out or args.no_dashboard:
        report = build_json_report(path, core, app, rsid, risk, fmt, extra)
        json_str = json.dumps(report, indent=2, ensure_ascii=False, default=str)
        if args.out:
            Path(args.out).write_text(json_str, encoding="utf-8")
            console.print(f"[green]Laporan JSON disimpan ke:[/green] {args.out}")
        if args.json or args.no_dashboard:
            print(json_str)


if __name__ == "__main__":
    main()
