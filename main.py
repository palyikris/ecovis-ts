# main.py
# -*- coding: utf-8 -*-
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import subprocess
import datetime
import threading
import sys
import queue
import re
from pathlib import Path
import tkinter as tk  # Listboxhoz
from tkinter import filedialog, messagebox
import os
import time
from typing import Optional, Dict, Any, List
from PIL import Image, ImageTk
import json
import shutil
import smtplib
from email.message import EmailMessage

# === Beállítások külön modulban ===
from settings import SETTINGS, save_settings, DEFAULT_SETTINGS, CONFIG_PATH

# ===========================
#  ÁLLANDÓK / SEGÉDFÜGGVÉNYEK
# ===========================

APP_TITLE = "Ecovis Timesheet Tool"

# --- a kódlista betöltéséhez/ALAPÉRTELMEZETT listához használjuk a generátor segédfüggvényeit
try:
    from generate_szamlamelleklet import (
        load_client_name_map,
        ORDERED_CODES_DEFAULT,
        remove_accents,
    )
except Exception:
    load_client_name_map = None
    ORDERED_CODES_DEFAULT = []

    def remove_accents(s: str) -> str:
        return s


# ⚙️ Config
MONTHS = [
    "Teljes év",
    "januar",
    "februar",
    "marcius",
    "aprilis",
    "majus",
    "junius",
    "julius",
    "augusztus",
    "szeptember",
    "oktober",
    "november",
    "december",
]
current_month = MONTHS[datetime.datetime.now().month]

# ---- Icons
ICON_RUNNING = "⏳"
ICON_FILE = "📄"
ICON_SHEET = "🗓️"
ICON_OK = "✅"
ICON_WARN = "⚠️"
ICON_ERR = "❌"

# ---- Thread-safe UI pipeline
ui_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
details_buffer: list[str] = []

# --- ügyfélkód-választás állapota
selected_client_codes: list[str] | None = None
all_client_codes_sorted: list[str] = []

# --- Utolsó futás ideje (dashboardon mutatjuk)
last_run_duration_s: Optional[float] = None

# ---- Regexek a log feldolgozásához
FILE_RE = re.compile(r"Feldolgozás:\s*(.+\.xlsx)", re.IGNORECASE)
SHEET_RE = re.compile(r"Sheet:\s*(.+)", re.IGNORECASE)
SKIP_RE = re.compile(r"Kihagyva", re.IGNORECASE)
DONE_FILE_RE = re.compile(r"^✅ Kész:\s*(.+\.xlsx)$")
SUMMARY_RE = re.compile(
    r"Run summary|Összesítés elkészült|Kész a formázott hibalista|Nincs hiba",
    re.IGNORECASE,
)
ERROR_RE = re.compile(r"(❌|hiba)", re.IGNORECASE)

# Lazított minta: bármely sor, amiben "Kész" és ".xlsx" szerepel
OUTPUT_LINE_RE = re.compile(r"Kész.*?\.xlsx", re.IGNORECASE)
FILEPATH_XLSX_RE = re.compile(
    r'([A-Za-z]:[\\/][^<>:"|?*\n\r]+\.xlsx|[^ \t\n\r<>:"|?*]+\.xlsx)'
)

# ===========================
#    FÁJLRÖGZÍTŐ SEGÉDEK
# ===========================


def post(level: str, msg: str):
    ui_queue.put((level, msg))


def limit_push(buffer: list[str], line: str, limit: int = 50):
    buffer.append(line.rstrip())
    if len(buffer) > limit:
        buffer.pop(0)


def ts_root() -> Path:
    p = Path(SETTINGS.get("ts_folder") or Path.cwd())
    return p if p.exists() else Path.cwd()


def output_root() -> Path:
    p = Path(SETTINGS.get("output_folder") or "")
    if not p:
        return ts_root()
    return p if p.exists() else ts_root()


def reports_root() -> Path:
    r = output_root() / "reports"
    r.mkdir(parents=True, exist_ok=True)
    return r


def backup_root() -> Optional[Path]:
    if not SETTINGS.get("backup_enabled"):
        return None
    p = Path(SETTINGS.get("backup_folder") or "")
    return p if p.exists() else None


def parse_and_emit(line: str, _title: str):
    text = line.strip()
    if not text:
        return
    limit_push(details_buffer, text)

    m = FILE_RE.search(text)
    if m:
        post("info", f"{ICON_FILE} Fájl: {m.group(1)}")
        return

    m = SHEET_RE.search(text)
    if m:
        post("info", f"{ICON_SHEET} Hónap lap: {m.group(1)}")
        return

    if SKIP_RE.search(text):
        post("warn", f"{ICON_WARN} {text}")
        return

    m = DONE_FILE_RE.search(text)
    if m:
        post("ok", f"{ICON_OK} Kész: {m.group(1)}")
        return

    if SUMMARY_RE.search(text):
        if "Nincs hiba" in text:
            post("ok", f"{ICON_OK} Nincs hiba a jelentésben")
        elif "Összesítés elkészült" in text:
            post("ok", f"{ICON_OK} Összesítés elkészült")
        elif "Kész a formázott hibalista" in text:
            post("warn", f"{ICON_WARN} Formázott hibalista elkészült")
        else:
            post("info", "📊 Futás összegzése kész")
        return

    if ERROR_RE.search(text):
        post("err", f"{ICON_ERR} {text}")
        return


# ---- Platformfüggetlen megnyitás
def open_path(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        post("err", f"{ICON_ERR} Megnyitási hiba: {e}")


def open_file(path: Path):
    if not path:
        return
    p = path if path.is_absolute() else (Path.cwd() / path)
    if not p.exists():
        post("err", f"{ICON_ERR} A fájl nem található: {p}")
        return
    open_path(p)


# ---- Fájlrendszer segédek (Dashboardhoz + progress összesítéshez)
def list_ts_files() -> list[Path]:
    root = ts_root()
    files = []
    for p in root.glob("*.xlsx"):
        name = p.name
        if name.startswith("~$"):
            continue
        if "TS" in name:
            files.append(p)
    return files


def latest_of(globs: list[str]) -> Optional[Path]:
    """Visszaadja a legutóbb módosított fájlt a megadott globok közül (TS mappában vagy Output mappában)."""
    candidates: list[Path] = []
    for base in [output_root(), ts_root()]:
        for g in globs:
            candidates.extend(base.glob(g))
    candidates = [c for c in candidates if not c.name.startswith("~$")]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def fmt_ts(ts: float) -> str:
    dt = datetime.datetime.fromtimestamp(ts)
    return dt.strftime("%Y-%m-%d %H:%M")


def maybe_beep():
    if SETTINGS.get("sound_enabled"):
        try:
            app.bell()
        except Exception:
            pass


# ===========================
#            RUNNER
# ===========================


def run_task(
    cmd: list[str],
    title_for_dialog: str,
    progressbar: tb.Progressbar,
    pb_counter_label: tb.Label,
    expects_output_file: bool = False,
    progressable: bool = True,  # ha True, i/N számlálót mutat
    expected_globs: list[str] | None = None,  # futás utáni fallback kereséshez
):
    def worker():
        global details_buffer, last_run_duration_s
        details_buffer = []

        post("info", f"{ICON_RUNNING} {title_for_dialog} elindult…")

        # progress init a főszálon
        total = len(list_ts_files()) if progressable else 0

        def init_progress():
            if progressable and total > 0:
                progressbar.configure(mode="determinate", maximum=total, value=0)
                pb_counter_label.config(text=f"0/{total}")
            else:
                progressbar.configure(mode="indeterminate")
                pb_counter_label.config(text="—")
                progressbar.start()

        app.after(0, init_progress)

        processed = 0
        output_path: Path | None = None

        def try_capture_output_path(line: str):
            nonlocal output_path
            if not expects_output_file:
                return
            if OUTPUT_LINE_RE.search(line):
                for m in FILEPATH_XLSX_RE.findall(line):
                    cand = Path(m)
                    # a script TS mappából fut
                    cand_abs = (ts_root() / cand) if not cand.is_absolute() else cand
                    output_path = cand_abs

        fs_started_at = time.time()  # fájlrendszer "óra" a fallbackhez
        t0 = time.perf_counter()
        try:
            # Subprocess-t a TS mappában futtatjuk
            proc = subprocess.Popen(
                cmd,
                cwd=str(ts_root()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert proc.stdout is not None
            for line in proc.stdout:
                parse_and_emit(line, title_for_dialog)
                try_capture_output_path(line)

                # determinált progress léptetés a „Feldolgozás:” sorokra
                if progressable and FILE_RE.search(line):
                    processed += 1

                    def step():
                        if progressbar["mode"] == "indeterminate":
                            progressbar.stop()
                            progressbar.configure(
                                mode="determinate", maximum=max(1, total), value=0
                            )
                        progressbar["value"] = min(
                            processed, int(progressbar["maximum"])
                        )
                        if total > 0:
                            pb_counter_label.config(
                                text=f"{min(processed, total)}/{total}"
                            )
                        else:
                            pb_counter_label.config(text=str(processed))

                    app.after(0, step)

            rc = proc.wait()
            last_run_duration_s = max(0.0, time.perf_counter() - t0)

            # ---- Fallback: ha nem találtunk fájlnevet a logból, de várunk kimenetet
            if rc == 0 and expects_output_file and output_path is None:
                globs = list(expected_globs or [])
                # végső tartalék: típusszerinti minta
                if (
                    "Számlamelléklet" in title_for_dialog
                    and "szamlamelleklet_*.xlsx" not in globs
                ):
                    globs.append("szamlamelleklet_*.xlsx")
                if (
                    "Összesített idők" in title_for_dialog
                    and "timesheet_summary_*.xlsx" not in globs
                ):
                    globs.append("timesheet_summary_*.xlsx")
                if (
                    "Párellenőrzés" in title_for_dialog
                    and "invalid_parok_*.xlsx" not in globs
                ):
                    globs.append("invalid_parok_*.xlsx")

                # TS vagy Output mappában keressük a legfrissebbet a futás óta
                def find_recent_by_globs(
                    globs_list: list[str], since_ts: float
                ) -> Optional[Path]:
                    candidates: list[Path] = []
                    for base in [output_root(), ts_root()]:
                        for g in globs_list:
                            candidates.extend(
                                p for p in base.glob(g) if not p.name.startswith("~$")
                            )
                    if not candidates:
                        return None
                    recent = [
                        c for c in candidates if c.stat().st_mtime >= since_ts - 1.0
                    ]
                    if not recent:
                        return None
                    return max(recent, key=lambda p: p.stat().st_mtime)

                guess = find_recent_by_globs(globs, fs_started_at)
                if guess:
                    output_path = guess
                    post("info", f"{ICON_OK} Kimeneti fájl: {output_path.name}")

            # Ha megvan a kimenet és van megadott output mappa: másoljuk oda és onnan használjuk
            if rc == 0 and expects_output_file and output_path is not None:
                dest_dir = output_root()
                if dest_dir and output_path.parent != dest_dir:
                    try:
                        dest_dir.mkdir(parents=True, exist_ok=True)
                        dest = dest_dir / output_path.name
                        shutil.copy2(str(output_path), str(dest))
                        output_path = dest
                    except Exception as e:
                        post(
                            "warn",
                            f"{ICON_WARN} Nem sikerült a kimenetet az output mappába másolni: {e}",
                        )

                # Backup (ha be van kapcsolva)
                bdir = backup_root()
                if bdir:
                    try:
                        bdir.mkdir(parents=True, exist_ok=True)
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        bname = output_path.stem + f"_{ts}" + output_path.suffix
                        shutil.copy2(str(output_path), str(bdir / bname))
                    except Exception as e:
                        post("warn", f"{ICON_WARN} Backup nem sikerült: {e}")

            def finish(success: bool):
                try:
                    if progressbar["mode"] == "indeterminate":
                        progressbar.stop()
                except Exception:
                    pass
                # státuszsor – időtartammal
                if success:
                    status_label.config(
                        text=f"✅ {title_for_dialog} — {last_run_duration_s:.1f}s"
                    )
                else:
                    status_label.config(
                        text=f"❌ {title_for_dialog} — hiba (rc={rc}) — {last_run_duration_s:.1f}s"
                    )
                maybe_beep()

                # Popup
                show_result_dialog(success, title_for_dialog, output_path)

                # Hiba esetén: automatikus „Részletek” felugró (log)
                if not success and SETTINGS.get("auto_open_details_on_error"):
                    show_details_window()

            if rc == 0:
                post("ok", f"{ICON_OK} {title_for_dialog} sikeres")
                app.after(0, lambda: finish(True))
            else:
                post("err", f"{ICON_ERR} {title_for_dialog} hibával zárult (kód: {rc})")
                app.after(0, lambda: finish(False))

        except Exception as e:
            last_run_duration_s = max(0.0, time.perf_counter() - t0)
            post("err", f"{ICON_ERR} Hiba: {e}")

            def fail():
                try:
                    if progressbar["mode"] == "indeterminate":
                        progressbar.stop()
                except Exception:
                    pass
                status_label.config(
                    text=f"❌ {title_for_dialog} — kivétel — {last_run_duration_s:.1f}s"
                )
                maybe_beep()
                show_result_dialog(False, title_for_dialog, None)
                if SETTINGS.get("auto_open_details_on_error"):
                    show_details_window()

            app.after(0, fail)

        # művelet vége: dashboard frissítése
        app.after(0, refresh_dashboard)

    threading.Thread(target=worker, daemon=True).start()


def show_result_dialog(success: bool, title: str, output_path: Path | None):
    # „Részletek” helyett: „Kész file megnyitása” + automatikus zárás, automatikus megnyitás opció
    dlg = tb.Toplevel(title="Eredmény")
    dlg.transient(app)
    dlg.grab_set()
    dlg.resizable(False, False)

    frm = tb.Frame(dlg, padding=30)
    frm.pack(fill=BOTH, expand=True)

    icon_lbl = tb.Label(
        frm, text=("🟢" if success else "🔴"), font=("Segoe UI Emoji", 48)
    )
    icon_lbl.pack()

    title_lbl = tb.Label(
        frm,
        text=("Sikeres művelet" if success else "Hiba történt"),
        font=("Segoe UI", 20, "bold"),
    )
    title_lbl.pack(pady=(10, 4))

    sub_lbl = tb.Label(
        frm, text=title, font=("Segoe UI", 12), bootstyle=SUCCESS if success else DANGER
    )
    sub_lbl.pack(pady=(0, 12))

    btns = tb.Frame(frm)
    btns.pack(pady=(10, 0), fill=X)

    def openable(p: Path | None) -> bool:
        if not p:
            return False
        p = p if p.is_absolute() else (Path.cwd() / p)
        return p.exists()

    can_open = success and openable(output_path)
    auto_open = SETTINGS.get("auto_open_output_on_success") and can_open

    if can_open and not auto_open:
        tb.Button(
            btns,
            text="Kész file megnyitása",
            bootstyle=PRIMARY,
            command=lambda p=output_path: (dlg.destroy(), open_file(p)),
        ).pack(side=LEFT, padx=(0, 8))

    # ha automatikus megnyitás be van kapcsolva, nyissuk is meg
    if auto_open:
        try:
            open_file(output_path)  # type: ignore[arg-type]
        except Exception:
            pass

    tb.Button(
        btns, text="OK", bootstyle=(SUCCESS if success else DANGER), command=dlg.destroy
    ).pack(side=RIGHT)

    # automatikus bezárás
    ac = int(SETTINGS.get("popup_autoclose_sec") or 0)
    if ac > 0:
        app.after(ac * 1000, lambda: (dlg.winfo_exists() and dlg.destroy()))


# ===========================
#  ÜGYFÉLKÓD-SELECTOR & LOG
# ===========================


def show_details_window():
    # shows up to 10 most relevant lines from details_buffer
    win = tb.Toplevel(title="Részletek")
    win.attributes("-fullscreen", True)
    win.transient(app)
    win.grab_set()

    info = tb.Label(
        win,
        text="Legfeljebb 10 releváns sor a futásból.",
        font=("Segoe UI", 10),
    )
    info.pack(padx=15, pady=(15, 5), anchor=W)

    text = tb.ScrolledText(win, wrap="word", font=("Consolas", 10))
    text.pack(fill=BOTH, expand=True, padx=15, pady=10)

    # filter for relevant-looking lines first
    relevant = []
    for ln in details_buffer:
        if any(
            key in ln
            for key in (
                "Feldolgozás:",
                "Sheet:",
                "Kész:",
                "Kihagyva",
                "Összesítés",
                "Nincs hiba",
                "Hiba",
                "Nem sikerült",
            )
        ):
            relevant.append(ln)

    # fallback: if not enough relevant, add recent tail
    if len(relevant) < 10:
        tail = [
            ln for ln in details_buffer[-20:]
        ]  # take last 20, will slice to 10 below
        # keep uniqueness while preserving order
        seen = set(relevant)
        for ln in tail:
            if ln not in seen:
                relevant.append(ln)
                seen.add(ln)

    # cap to 10 lines
    relevant = relevant[:10] if relevant else ["(Nincs megjeleníthető részlet.)"]

    text.insert("1.0", "\n".join(relevant))
    text.configure(state="disabled")


def open_client_code_selector():
    global all_client_codes_sorted, selected_client_codes
    try:
        if load_client_name_map is None:
            raise RuntimeError("A generate_szamlamelleklet modul nem érhető el.")
        name_map = load_client_name_map()
        all_client_codes_sorted = sorted(name_map.keys(), key=remove_accents)
    except Exception as e:
        post("err", f"{ICON_ERR} Hiba az ügyfélkódok betöltésekor: {e}")
        return

    # preselect for settings: default_client_codes or ORDERED_CODES_DEFAULT
    defaults_from_settings = [
        c.strip()
        for c in (SETTINGS.get("default_client_codes") or "").split(",")
        if c.strip()
    ]
    defaults_base = defaults_from_settings or [c for c in ORDERED_CODES_DEFAULT]

    preselect = (
        selected_client_codes
        if (selected_client_codes is not None and len(selected_client_codes) > 0)
        else [c for c in defaults_base if c in all_client_codes_sorted]
    )

    dlg = tb.Toplevel(title="Ügyfélkódok kiválasztása")
    dlg.transient(app)
    dlg.grab_set()
    dlg.geometry("520x600")

    frm = tb.Frame(dlg, padding=15)
    frm.pack(fill=BOTH, expand=True)

    tb.Label(
        frm,
        text="Válaszd ki, mely ügyfélkódokhoz készüljön számlamelléklet.",
        font=("Segoe UI", 11),
    ).pack(anchor=W, pady=(0, 10))

    btnbar = tb.Frame(frm)
    btnbar.pack(fill=X, pady=(0, 8))

    def do_select_all():
        listbox.selection_set(0, tk.END)
        update_count()

    def do_clear_all():
        listbox.selection_clear(0, tk.END)
        update_count()

    def do_defaults():
        listbox.selection_clear(0, tk.END)
        for i, val in enumerate(all_client_codes_sorted):
            if val in defaults_base:
                listbox.selection_set(i)
        update_count()

    tb.Button(
        btnbar, text="Kijelöl mind", bootstyle=SECONDARY, command=do_select_all
    ).pack(side=LEFT, padx=(0, 6))
    tb.Button(
        btnbar, text="Töröl mind", bootstyle=SECONDARY, command=do_clear_all
    ).pack(side=LEFT, padx=6)
    tb.Button(btnbar, text="Alapértelmezett", bootstyle=INFO, command=do_defaults).pack(
        side=LEFT, padx=6
    )

    cnt_lbl = tb.Label(btnbar, text="", bootstyle=SECONDARY)
    cnt_lbl.pack(side=RIGHT)

    listbox = tk.Listbox(frm, selectmode="extended", activestyle="dotbox")
    vs = tb.Scrollbar(frm, orient="vertical", command=listbox.yview)
    listbox.configure(yscrollcommand=vs.set, font=("Segoe UI", 10))
    listbox.pack(side=LEFT, fill=BOTH, expand=True)
    vs.pack(side=RIGHT, fill=Y)

    for code in all_client_codes_sorted:
        listbox.insert(tk.END, code)
    for i, val in enumerate(all_client_codes_sorted):
        if val in preselect:
            listbox.selection_set(i)

    def update_count(*_):
        sel = len(listbox.curselection())
        cnt_lbl.config(text=f"{sel} / {len(all_client_codes_sorted)} kijelölve")

    listbox.bind("<<ListboxSelect>>", update_count)
    update_count()

    act = tb.Frame(frm)
    act.pack(fill=X, pady=(10, 0))

    def on_save():
        global selected_client_codes
        idx = listbox.curselection()
        selected_client_codes = [listbox.get(i) for i in idx]
        post(
            "info",
            f"{ICON_OK} Ügyfélkód-választás elmentve ({len(selected_client_codes)} kód).",
        )
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    tb.Button(act, text="Mentés", bootstyle=SUCCESS, command=on_save).pack(
        side=RIGHT, padx=(6, 0)
    )
    tb.Button(act, text="Mégse", bootstyle=SECONDARY, command=on_cancel).pack(
        side=RIGHT
    )


# ===========================
#           HANDLEREK
# ===========================


def update_dropdowns():
    run_task(
        [sys.executable, "update_dropdowns.py"],
        "Legördülő elemek frissítése",
        dropdown_progress,
        dropdown_info,
        expects_output_file=False,
        progressable=True,
    )


def aggregate_hours():
    selected_month = month_var.get() or current_month
    run_task(
        [sys.executable, "timesheet_summary.py", selected_month],
        f"Összesített idők — hónap: {selected_month}",
        agg_progress,
        agg_info,
        expects_output_file=True,
        progressable=True,
        expected_globs=[
            f"timesheet_summary_{selected_month}.xlsx",
            "timesheet_summary_*.xlsx",
        ],
    )


def generate_szamlamelleklet():
    selected_month = month_var.get() or current_month
    if selected_client_codes and len(selected_client_codes) > 0:
        cmd = [
            sys.executable,
            "generate_szamlamelleklet.py",
            selected_month,
            *selected_client_codes,
        ]
        suffix = f" ({len(selected_client_codes)} ügyfélkód)"
    else:
        # ha vannak default kódok a settingsben, adjuk hozzá
        defaults = [
            c.strip()
            for c in (SETTINGS.get("default_client_codes") or "").split(",")
            if c.strip()
        ]
        cmd = (
            [sys.executable, "generate_szamlamelleklet.py", selected_month, *defaults]
            if defaults
            else [sys.executable, "generate_szamlamelleklet.py", selected_month]
        )
        suffix = (
            " (alapértelmezett ügyfélkód lista)"
            if not defaults
            else f" ({len(defaults)} alapértelmezett kód)"
        )

    run_task(
        cmd,
        f"Számlamelléklet — hónap: {selected_month}{suffix}",
        szamla_progress,
        szamla_info,
        expects_output_file=True,
        progressable=True,
        expected_globs=[
            f"szamlamelleklet_{selected_month}.xlsx",
            "szamlamelleklet_*.xlsx",
        ],
    )


def validate_pairs():
    selected_month = month_var.get() or current_month
    run_task(
        [sys.executable, "validate_pairs.py", selected_month],
        f"Ügyfélkód–Projekt párellenőrzés — hónap: {selected_month}",
        pairs_progress,
        pairs_info,
        expects_output_file=True,
        progressable=True,
        expected_globs=[f"invalid_parok_{selected_month}.xlsx", "invalid_parok_*.xlsx"],
    )


# ÚJ: TS reset handler
def reset_timesheets():
    """TS fájlok archiválása és üres példányok létrehozása (reset_timesheets.py)."""
    ts_files = list_ts_files()
    if not ts_files:
        messagebox.showinfo("TS reset", "Nem található TS fájl a mappában.")
        return

    msg = (
        f"{len(ts_files)} TS fájl archiválásra kerül az archív mappába,\n"
        f"és ugyanennyi üres példány jön létre azonos névvel és beállításokkal.\n\n"
        f"Folytatod?"
    )
    if not messagebox.askyesno("TS reset — megerősítés", msg):
        return

    run_task(
        [sys.executable, "reset_timesheets.py"],
        "TS reset — archiválás és üres fájlok létrehozása",
        reset_progress,
        reset_info,
        expects_output_file=False,
        progressable=True,
    )


# ---------------------------
#  📆 Havi zárás (pipeline)
# ---------------------------


def month_close_pipeline():
    """Egygombos havi zárás: update -> summary -> validate -> invoice"""
    selected_month = month_var.get() or current_month

    # Pre-flight
    ts_files = list_ts_files()
    if not ts_files:
        messagebox.showwarning("Havi zárás", "Nem található TS fájl a TS mappában.")
        return

    # Ügyfélkódok a számlához
    if selected_client_codes and len(selected_client_codes) > 0:
        codes_for_invoice = selected_client_codes[:]
    else:
        defaults = [
            c.strip()
            for c in (SETTINGS.get("default_client_codes") or "").split(",")
            if c.strip()
        ]
        codes_for_invoice = defaults

    dlg = tb.Toplevel(title="📆 Havi zárás")
    dlg.transient(app)
    dlg.grab_set()
    dlg.resizable(False, False)

    frm = tb.Frame(dlg, padding=20)
    frm.pack(fill=BOTH, expand=True)

    title = tb.Label(
        frm, text=f"📆 Havi zárás — {selected_month}", font=("Segoe UI", 14, "bold")
    )
    title.pack(anchor=W)

    step_lbl = tb.Label(frm, text="Előkészítés…", font=("Segoe UI", 11))
    step_lbl.pack(anchor=W, pady=(8, 6))

    pbar = tb.Progressbar(frm, mode="determinate", maximum=4, value=0, bootstyle=INFO)
    pbar.pack(fill=X)

    log_hint = tb.Label(
        frm, text="A részletek a naplóban (alul) is követhetők.", bootstyle=SECONDARY
    )
    log_hint.pack(anchor=W, pady=(8, 0))

    results: Dict[str, Optional[Path]] = {
        "summary": None,
        "validate": None,
        "invoice": None,
    }

    def run_step(
        cmd: List[str], title_for_dialog: str, expected_globs: List[str] | None
    ) -> tuple[bool, Optional[Path]]:
        nonlocal step_lbl, pbar
        app.after(0, lambda: step_lbl.config(text=title_for_dialog))
        post("info", f"{ICON_RUNNING} {title_for_dialog}…")

        output_path: Optional[Path] = None
        fs_started_at = time.time()

        def try_capture_output_path(line: str):
            nonlocal output_path
            if OUTPUT_LINE_RE.search(line):
                for m in FILEPATH_XLSX_RE.findall(line):
                    cand = Path(m)
                    cand_abs = (ts_root() / cand) if not cand.is_absolute() else cand
                    output_path = cand_abs

        try:
            proc = subprocess.Popen(
                cmd,
                cwd=str(ts_root()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert proc.stdout is not None
            for line in proc.stdout:
                parse_and_emit(line, title_for_dialog)
                try_capture_output_path(line)
            rc = proc.wait()
            if rc != 0:
                post("err", f"{ICON_ERR} Hiba: {title_for_dialog} (rc={rc})")
                return False, None

            if output_path is None and expected_globs:
                candidates: list[Path] = []
                for base in [output_root(), ts_root()]:
                    for g in expected_globs:
                        candidates.extend(
                            p for p in base.glob(g) if not p.name.startswith("~$")
                        )
                recent = [
                    c for c in candidates if c.stat().st_mtime >= fs_started_at - 1.0
                ] or candidates
                if recent:
                    output_path = max(recent, key=lambda p: p.stat().st_mtime)

            if output_path is not None:
                dest_dir = output_root()
                if dest_dir and output_path.parent != dest_dir:
                    try:
                        dest_dir.mkdir(parents=True, exist_ok=True)
                        dest = dest_dir / output_path.name
                        shutil.copy2(str(output_path), str(dest))
                        output_path = dest
                    except Exception as e:
                        post(
                            "warn",
                            f"{ICON_WARN} Nem sikerült az output mappába másolni: {e}",
                        )

            post("ok", f"{ICON_OK} Kész: {title_for_dialog}")
            return True, output_path
        except Exception as e:
            post("err", f"{ICON_ERR} Kivétel: {title_for_dialog}: {e}")
            return False, None

    def worker():
        ok, _ = run_step(
            [sys.executable, "update_dropdowns.py"], "Legördülők frissítése", None
        )
        app.after(0, lambda: pbar.configure(value=1))
        if not ok:
            app.after(
                0,
                lambda: (
                    dlg.destroy(),
                    messagebox.showerror(
                        "Havi zárás", "Megakadt: Legördülők frissítése."
                    ),
                ),
            )
            return

        month = month_var.get() or current_month
        ok, summary_file = run_step(
            [sys.executable, "timesheet_summary.py", month],
            f"Összesített idők — hónap: {month}",
            [f"timesheet_summary_{month}.xlsx", "timesheet_summary_*.xlsx"],
        )
        results["summary"] = summary_file
        app.after(0, lambda: pbar.configure(value=2))
        if not ok:
            app.after(
                0,
                lambda: (
                    dlg.destroy(),
                    messagebox.showerror("Havi zárás", "Megakadt: Összesített idők."),
                ),
            )
            return

        ok, validate_file = run_step(
            [sys.executable, "validate_pairs.py", month],
            f"Ügyfélkód–Projekt párellenőrzés — hónap: {month}",
            [f"invalid_parok_{month}.xlsx", "invalid_parok_*.xlsx"],
        )
        results["validate"] = validate_file
        app.after(0, lambda: pbar.configure(value=3))
        if not ok:
            app.after(
                0,
                lambda: (
                    dlg.destroy(),
                    messagebox.showerror("Havi zárás", "Megakadt: Párellenőrzés."),
                ),
            )
            return

        cmd = [
            sys.executable,
            "generate_szamlamelleklet.py",
            month,
            *([c for c in (selected_client_codes or [])] or []),
        ]
        ok, invoice_file = run_step(
            cmd,
            f"Számlamelléklet — hónap: {month}",
            [f"szamlamelleklet_{month}.xlsx", "szamlamelleklet_*.xlsx"],
        )
        results["invoice"] = invoice_file
        app.after(0, lambda: pbar.configure(value=4))
        if not ok:
            app.after(
                0,
                lambda: (
                    dlg.destroy(),
                    messagebox.showerror("Havi zárás", "Megakadt: Számlamelléklet."),
                ),
            )
            return

        def show_summary():
            dlg.destroy()
            sdlg = tb.Toplevel(title="Havi zárás — összegzés")
            sdlg.transient(app)
            sdlg.grab_set()
            sdlg.resizable(False, False)

            sfrm = tb.Frame(sdlg, padding=20)
            sfrm.pack(fill=BOTH, expand=True)

            tb.Label(
                sfrm, text="✅ Havi zárás kész", font=("Segoe UI", 14, "bold")
            ).pack(anchor=W)
            tb.Label(sfrm, text=f"Hónap: {month}", bootstyle=SECONDARY).pack(
                anchor=W, pady=(0, 8)
            )

            def row(label: str, path: Optional[Path]):
                r = tb.Frame(sfrm)
                r.pack(fill=X, pady=2)
                tb.Label(r, text=label + ":", width=22, anchor=W).pack(side=LEFT)
                tb.Label(r, text=(path.name if path else "—")).pack(
                    side=LEFT, padx=(6, 8)
                )
                if path and path.exists():
                    tb.Button(
                        r,
                        text="Megnyitás",
                        bootstyle=PRIMARY,
                        command=lambda p=path: open_file(p),
                    ).pack(side=LEFT)

            row("Összesítés", results["summary"])
            row("Párellenőrzés", results["validate"])
            row("Számlamelléklet", results["invoice"])

            btns = tb.Frame(sfrm)
            btns.pack(fill=X, pady=(10, 0))
            tb.Button(
                btns,
                text="📂 Mappa megnyitása",
                bootstyle=INFO,
                command=lambda: open_path(output_root()),
            ).pack(side=LEFT)
            tb.Button(btns, text="OK", bootstyle=SUCCESS, command=sdlg.destroy).pack(
                side=RIGHT
            )

        app.after(0, show_summary)
        app.after(0, refresh_dashboard)

    threading.Thread(target=worker, daemon=True).start()


# ===========================
#         DASHBOARD
# ===========================


def refresh_dashboard():
    ts_files = list_ts_files()
    ts_val.config(text=str(len(ts_files)))
    ts_sub.config(text=("Nincs TS fájl" if not ts_files else str(ts_root())))

    latest_summary = latest_of(["timesheet_summary_*.xlsx"])
    if latest_summary:
        m = latest_summary.stat().st_mtime
        sum_val.config(text=ellipsize_middle(latest_summary.name, 24))
        sum_sub.config(text=f"Módosítva: {fmt_ts(m)}")
    else:
        sum_val.config(text="—")
        sum_sub.config(text="Még nincs összesítés")

    latest_invoice = latest_of(["szamlamelleklet_*.xlsx"])
    if latest_invoice:
        m = latest_invoice.stat().st_mtime
        inv_val.config(text=ellipsize_middle(latest_invoice.name, 24))
        inv_sub.config(text=f"Módosítva: {fmt_ts(m)}")
    else:
        inv_val.config(text="—")
        inv_sub.config(text="Még nincs számlamelléklet")

    latest_invalid = latest_of(["invalid_parok_*.xlsx"])
    if latest_invalid:
        m = latest_invalid.stat().st_mtime
        val_val.config(text=ellipsize_middle(latest_invalid.name, 24))
        val_sub.config(text=f"Módosítva: {fmt_ts(m)}")
    else:
        val_val.config(text="—")
        val_sub.config(text="Még nincs hibalista")

    if last_run_duration_s is not None:
        run_val.config(text=f"{last_run_duration_s:.1f} s")
        run_sub.config(text="Legutóbbi művelet")
    else:
        run_val.config(text="—")
        run_sub.config(text="Még nincs futás")


# ===========================
#         HETI RIPORT
# ===========================


def parse_hhmm(s: str) -> Optional[datetime.time]:
    try:
        h, m = s.strip().split(":")
        hh, mm = int(h), int(m)
        if 0 <= hh < 24 and 0 <= mm < 60:
            return datetime.time(hour=hh, minute=mm)
    except Exception:
        pass
    return None


def _email_list(s: str) -> List[str]:
    return [x.strip() for x in (s or "").split(",") if x.strip()]


# --- EMAIL KÜLDÉS: részletes hibával tér vissza (ok: bool, err: str) ----
def send_email_outlook(
    subject: str, body: str, to_list: list[str], attachments: list[Path]
) -> tuple[bool, str]:
    try:
        import win32com.client as win32  # type: ignore
        import pythoncom  # <- COM szál-inicializálás
    except Exception as e:
        return False, f"Outlook COM nem elérhető: {e!r}"

    try:
        pythoncom.CoInitialize()  # **KELL** háttérszálban
        try:
            # EnsureDispatch megbízhatóbb, mint a sima Dispatch
            outlook = win32.gencache.EnsureDispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "; ".join(to_list)
            mail.Subject = subject
            mail.Body = body
            for p in attachments:
                if p and p.exists() and not p.name.startswith("~$"):
                    mail.Attachments.Add(str(p))
            mail.Send()
            return True, ""
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        return False, f"Outlook küldési hiba: {e!r}"


def send_email_smtp(
    subject: str,
    body: str,
    to_list: List[str],
    attachments: List[Path],
    host: str,
    port: int,
    use_tls: bool,
    user: str,
    pwd: str,
) -> tuple[bool, str]:
    if not host or not port:
        return False, "SMTP host/port hiányzik."
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = user or "noreply@example.com"
        msg["To"] = ", ".join(to_list)
        msg.set_content(body)
        for p in attachments:
            if p and p.exists() and not p.name.startswith("~$"):
                data = p.read_bytes()
                msg.add_attachment(
                    data,
                    maintype="application",
                    subtype="octet-stream",
                    filename=p.name,
                )
        with smtplib.SMTP(host, port, timeout=30) as s:
            if use_tls:
                s.starttls()
            if user:
                s.login(user, pwd)
            s.send_message(msg)
        return True, ""
    except Exception as e:
        return False, f"SMTP küldési hiba: {e!r}"


def send_email(
    subject: str,
    body: str,
    to_list: List[str],
    attachments: List[Path],
    method: str,
    smtp_cfg: dict,
) -> tuple[bool, str]:
    m = (method or "outlook").lower()
    if m == "outlook":
        ok, err = send_email_outlook(subject, body, to_list, attachments)
        if ok:
            return True, ""
        # fallback SMTP, ha meg van adva
        if smtp_cfg.get("host"):
            return send_email_smtp(
                subject,
                body,
                to_list,
                attachments,
                smtp_cfg.get("host", ""),
                int(smtp_cfg.get("port", 587)),
                bool(smtp_cfg.get("tls", True)),
                smtp_cfg.get("user", ""),
                smtp_cfg.get("pwd", ""),
            )
        return False, err or "Outlook nem elérhető, és SMTP sincs beállítva."
    else:
        return send_email_smtp(
            subject,
            body,
            to_list,
            attachments,
            smtp_cfg.get("host", ""),
            int(smtp_cfg.get("port", 587)),
            bool(smtp_cfg.get("tls", True)),
            smtp_cfg.get("user", ""),
            smtp_cfg.get("pwd", ""),
        )


def weekly_report_now():
    """Kézi indítás: heti riport (summary + validate), email küldés csatolmányokkal, számlamelléklet nélkül."""

    def worker():
        month = month_var.get() or current_month
        post("info", "📧 Heti riport: indítás…")

        # 1) Összesített idők
        try:
            proc = subprocess.Popen(
                [sys.executable, "timesheet_summary.py", month],
                cwd=str(ts_root()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            out = proc.communicate()[0]
            for line in out.splitlines():
                parse_and_emit(line, "Riport: Összesített idők")
        except Exception as e:
            post("err", f"{ICON_ERR} Riport: összesítés hiba: {e}")

        summary = latest_of(
            [f"timesheet_summary_{month}.xlsx", "timesheet_summary_*.xlsx"]
        )

        # 2) Párellenőrzés
        try:
            proc = subprocess.Popen(
                [sys.executable, "validate_pairs.py", month],
                cwd=str(ts_root()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            out = proc.communicate()[0]
            for line in out.splitlines():
                parse_and_emit(line, "Riport: Párellenőrzés")
        except Exception as e:
            post("err", f"{ICON_ERR} Riport: párellenőrzés hiba: {e}")

        invalid = latest_of([f"invalid_parok_{month}.xlsx", "invalid_parok_*.xlsx"])

        # Másolás az output mappába és csatolmányok listája
        attachments: List[Path] = []
        for p in [summary, invalid]:
            if p:
                try:
                    dest_dir = output_root()
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest = dest_dir / p.name
                    if p.resolve() != dest.resolve():
                        shutil.copy2(str(p), str(dest))
                        p = dest
                except Exception:
                    pass
            if p and p.exists() and not p.name.startswith("~$"):
                attachments.append(p)

        # TXT összefoglaló
        rep_dir = reports_root()
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
        txt = rep_dir / f"weekly_report_{ts}.txt"
        lines = [
            f"Ecovis heti riport — {datetime.datetime.now():%Y-%m-%d %H:%M}",
            f"Hónap: {month}",
            "",
            f"Összesítés: {summary.name if summary else '—'}",
            f"Párellenőrzés: {invalid.name if invalid else '—'}",
            "",
            "Megjegyzés: A heti riport nem tartalmaz számlamellékletet.",
        ]
        try:
            txt.write_text("\n".join(lines), encoding="utf-8")
            attachments.insert(0, txt)
        except Exception:
            pass

        # (email küldés része továbbra is kikommentezve – jelenlegi viselkedés megőrizve)
        refresh_dashboard()

    threading.Thread(target=worker, daemon=True).start()


def weekly_report_tick():
    """Heti ütemező: ha engedélyezve van és eljött a nap/idő, küldi a riportot (heti 1x)."""
    try:
        if SETTINGS.get("weekly_report_enabled"):
            wd = int(SETTINGS.get("weekly_report_weekday", 0))  # 0=Hétfő
            t = parse_hhmm(str(SETTINGS.get("weekly_report_time", "08:30")))
            now = datetime.datetime.now()
            if t:
                if now.weekday() == wd:
                    target_dt = datetime.datetime.combine(now.date(), t)
                    week_key = f"{now.isocalendar().year}-{now.isocalendar().week}"
                    if (
                        now >= target_dt
                        and SETTINGS.get("last_weekly_report_key") != week_key
                    ):
                        weekly_report_now()
                        SETTINGS["last_weekly_report_key"] = week_key
                        save_settings(SETTINGS)
    except Exception:
        pass
    finally:
        app.after(60_000, weekly_report_tick)  # percenként ellenőriz


# ===========================
#           UI PUMP
# ===========================


def pump_ui():
    try:
        while True:
            level, msg = ui_queue.get_nowait()
            if level == "ok":
                style = "success"
            elif level == "warn":
                style = "warning"
            elif level == "err":
                style = "danger"
            else:
                style = "info"
            log_list.insert("", "end", values=(msg,), tags=(style,))
            log_list.see(log_list.get_children()[-1])
    except queue.Empty:
        pass
    app.after(100, pump_ui)


# ===========================
#          BEÁLLÍTÁSOK TAB
# ===========================


def apply_theme(new_theme: str):
    try:
        app.style.theme_use(new_theme)
        SETTINGS["theme"] = new_theme
        save_settings(SETTINGS)
    except Exception as e:
        messagebox.showerror("Téma", f"Nem sikerült alkalmazni: {e}")


def apply_scale(pct: int):
    try:
        scale = max(50, min(200, pct)) / 100.0
        app.tk.call("tk", "scaling", scale)
        SETTINGS["ui_scale_pct"] = int(pct)
        save_settings(SETTINGS)
    except Exception as e:
        messagebox.showerror("Skála", f"Nem sikerült alkalmazni: {e}")


def choose_dir(field: tk.Entry):
    path = filedialog.askdirectory(initialdir=ts_root())
    if path:
        field.delete(0, tk.END)
        field.insert(0, path)

from generate_szamlamelleklet import load_active_clients

def _load_all_client_codes_sorted() -> List[str]:
    """Összes ismert ügyfélkód (ha elérhető), abc szerint ékezetlenítve."""
    try:
        if load_client_name_map is None:
            return []
        name_map = load_client_name_map()
        active_set = set(load_active_clients())
        name_map = {k: v for k, v in name_map.items() if k in active_set}
        return sorted(name_map.keys(), key=remove_accents)
    except Exception:
        return []


def _initial_default_codes_for_settings_ui(all_codes: List[str]) -> List[str]:
    """A settings UI induló listája:
    - ha van mentett lista -> az
    - különben a hardcoded ORDERED_CODES_DEFAULT
    (mindkettőt a létező kódokra szűrjük, ha van all_codes)"""
    saved = [
        c.strip()
        for c in (SETTINGS.get("default_client_codes") or "").split(",")
        if c.strip()
    ]
    base = saved if saved else list(ORDERED_CODES_DEFAULT)
    if not all_codes:
        # nincs master lista – fogadjuk el mindet
        # (comboboxba kézzel is írhat a user)
        return list(dict.fromkeys(base))  # unique, order keep
    # csak a létezőket hagyjuk
    filtered = [c for c in base if c in all_codes]
    # ha mentett listában semmi nem érvényes, de van hardcoded -> próbáld azzal
    if (not filtered) and (not saved) and ORDERED_CODES_DEFAULT:
        filtered = [c for c in ORDERED_CODES_DEFAULT if c in all_codes]
    return list(dict.fromkeys(filtered))  # unique, order keep


def settings_save():
    SETTINGS["language"] = lang_var.get()
    SETTINGS["ts_folder"] = ts_folder_var.get().strip() or str(Path.cwd())
    SETTINGS["output_folder"] = output_folder_var.get().strip()
    SETTINGS["backup_enabled"] = backup_enabled_var.get()
    SETTINGS["backup_folder"] = backup_folder_var.get().strip()
    # default ügyfélkódok: listbox tartalma -> csv
    codes = default_codes_listbox.get(0, tk.END)
    SETTINGS["default_client_codes"] = ",".join(codes)
    SETTINGS["remember_last_selection"] = remember_last_var.get()
    SETTINGS["auto_open_output_on_success"] = auto_open_output_var.get()
    SETTINGS["auto_open_details_on_error"] = auto_open_details_var.get()
    try:
        SETTINGS["popup_autoclose_sec"] = int(popup_autoclose_var.get())
    except Exception:
        SETTINGS["popup_autoclose_sec"] = 0
    SETTINGS["sound_enabled"] = sound_var.get()

    # napi emlékeztető
    SETTINGS["daily_reminder_enabled"] = daily_reminder_enabled_var.get()
    SETTINGS["daily_reminder_time"] = daily_reminder_time_var.get().strip() or "18:00"

    save_settings(SETTINGS)
    refresh_dashboard()
    messagebox.showinfo(
        "Beállítások", "Mentve. (Egyes változások csak újraindítás után teljesek.)"
    )


# ===========================
#             GUI
# ===========================

# 🌟 GUI setup
app = tb.Window(themename=SETTINGS.get("theme", "minty"))
app.title(APP_TITLE)
app.state("zoomed")

# skála
try:
    app.tk.call("tk", "scaling", (SETTINGS.get("ui_scale_pct", 100) / 100.0))
except Exception:
    pass

# Notebook (Főoldal + Beállítások)
notebook = tb.Notebook(app, bootstyle=PRIMARY)
notebook.pack(fill=BOTH, expand=True)

# ----- FŐOLDAL TAB
home_tab = tb.Frame(notebook)
notebook.add(home_tab, text="Főoldal")

content_frame = tb.Frame(home_tab, padding=50)
content_frame.pack(fill=BOTH, expand=True)

# --- Logo betöltése
try:
    logo_img = Image.open("ecovis_logo.png")
    logo_img = logo_img.resize((400, 60), Image.LANCZOS)  # szépen átméretezve
    logo_tk = ImageTk.PhotoImage(logo_img)
    logo_label = tb.Label(content_frame, image=logo_tk)
    logo_label.image = logo_tk  # referenciát tartani kell!
    logo_label.pack(pady=(0, 20))
except Exception as e:
    print(f"Nem sikerült betölteni a logót: {e}")

title_label = tb.Label(
    content_frame,
    text="📊 Ecovis Timesheet Management Tool",
    font=("Segoe UI", 28, "bold"),
)
title_label.pack(pady=(10, 40))

# === DASHBOARD SÁV ===========================================================


def ellipsize_middle(s: str, max_len: int = 24) -> str:
    s = str(s or "")
    if len(s) <= max_len:
        return s
    keep = max_len - 1
    left = keep // 2
    right = keep - left
    return s[:left] + "… " + s[-right:]


dash = tb.Frame(content_frame)
dash.pack(fill=X, pady=(0, 16))

# 5 egyforma szélességű oszlop
for i in range(5):
    dash.grid_columnconfigure(i, weight=1, uniform="dash")
dash.grid_rowconfigure(0, weight=1)


def mk_card_grid(parent, title: str, col: int):
    card = tb.Labelframe(parent, text=title, padding=12, bootstyle=SECONDARY)
    card.grid(row=0, column=col, sticky="nsew", padx=6)

    top = tb.Frame(card)
    top.pack(fill=X)
    val = tb.Label(top, text="—", font=("Segoe UI", 18, "bold"), anchor="w")
    val.pack(fill=X)

    # wraplength-et később dinamikusan állítjuk (resize-ra)
    sub = tb.Label(
        card,
        text="",
        font=("Segoe UI", 9),
        bootstyle=SECONDARY,
        anchor="w",
        justify="left",
    )
    sub.pack(fill=X, pady=(6, 0))
    return card, val, sub


card_ts, ts_val, ts_sub = mk_card_grid(dash, "📂 TS fájlok száma", 0)
card_sum, sum_val, sum_sub = mk_card_grid(dash, "📊 Legutóbbi összesítés", 1)
card_inv, inv_val, inv_sub = mk_card_grid(dash, "🧾 Legutóbbi számlamelléklet", 2)
card_val, val_val, val_sub = mk_card_grid(dash, "⚠️ Legutóbbi hibalista", 3)
card_run, run_val, run_sub = mk_card_grid(dash, "⏱️ Utolsó futás", 4)


# wraplength dinamikus frissítése, hogy tartalmak ne tolódjanak szét
def _update_card_wrap(event=None):
    try:
        # teljes szélesség / 5 oszlop - belső margók (~40px)
        col_w = max(160, (dash.winfo_width() // 5) - 40)
        for lbl in (ts_sub, sum_sub, inv_sub, val_sub, run_sub):
            lbl.configure(wraplength=col_w)
    except Exception:
        pass


dash.bind("<Configure>", _update_card_wrap)

# ============================================================================

# Actions
btns = tb.Frame(content_frame)
btns.pack(fill=X, pady=(0, 10))

update_btn = tb.Button(
    btns, text="🔄 Legördülők frissítése", bootstyle=SUCCESS, command=update_dropdowns
)
update_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

month_var = tk.StringVar()
month_dropdown = tb.Combobox(
    btns,
    textvariable=month_var,
    values=MONTHS,
    font=("Segoe UI", 12),
    state="readonly",
    width=16,
    bootstyle=PRIMARY,
)
month_dropdown.set("")
month_dropdown.pack(side=LEFT, padx=(16, 6))

agg_btn = tb.Button(
    btns, text="📊 Összesített Idők", bootstyle=INFO, command=aggregate_hours
)
agg_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

pairs_btn = tb.Button(
    btns, text="🧪 Párellenőrzés", bootstyle=WARNING, command=validate_pairs
)
pairs_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

codes_btn = tb.Button(
    btns, text="👥 Ügyfélkódok…", bootstyle=SECONDARY, command=open_client_code_selector
)
codes_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

szamla_btn = tb.Button(
    btns,
    text="🧾 Számlamelléklet",
    bootstyle=SECONDARY,
    command=generate_szamlamelleklet,
)
szamla_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

# ÚJ: 📆 Havi zárás gomb
month_close_btn = tb.Button(
    btns,
    text="📆 Havi zárás",
    bootstyle="outline-primary",
    command=month_close_pipeline,
)
month_close_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

# ÚJ: 📧 Riport küldése most
report_now_btn = tb.Button(
    btns,
    text="📧 Riport küldése most",
    bootstyle="outline-info",
    command=weekly_report_now,
)
report_now_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

# ÚJ: 🧹 TS reset gomb
reset_btn = tb.Button(
    btns,
    text="🧹 TS reset",
    bootstyle="outline-danger",
    command=reset_timesheets,
)
reset_btn.pack(side=LEFT, padx=6, ipadx=10, ipady=6)

# Progress bars row + per-bar számláló címkék
pb_row = tb.Frame(content_frame)
pb_row.pack(fill=X, pady=(0, 15))


def mk_pb(parent, style):
    frm = tb.Frame(parent)
    frm.pack(side=LEFT, padx=6, fill=X, expand=True)
    bar = tb.Progressbar(frm, mode="indeterminate", bootstyle=style)
    bar.pack(fill=X)
    lbl = tb.Label(frm, text="—", bootstyle=SECONDARY, font=("Segoe UI", 9))
    lbl.pack(anchor=E, pady=(2, 0))
    return bar, lbl


dropdown_progress, dropdown_info = mk_pb(pb_row, INFO)
agg_progress, agg_info = mk_pb(pb_row, SUCCESS)
pairs_progress, pairs_info = mk_pb(pb_row, WARNING)
szamla_progress, szamla_info = mk_pb(pb_row, SECONDARY)
# ÚJ: TS reset progress bar
reset_progress, reset_info = mk_pb(pb_row, DANGER)

# Log panel
log_card = tb.Labelframe(
    content_frame, text="Futás közben történt események", padding=10
)
log_card.pack(fill=BOTH, expand=True)

log_list = tb.Treeview(
    log_card, columns=("msg",), show="headings", height=14, bootstyle=INFO
)
log_list.heading("msg", text="Esemény")
log_list.column("msg", width=1200, anchor=W)
log_list.pack(fill=BOTH, expand=True)

log_list.tag_configure("success", foreground="#0f5132")
log_list.tag_configure("warning", foreground="#664d03")
log_list.tag_configure("danger", foreground="#842029")
log_list.tag_configure("info", foreground="#084298")

status_label = tb.Label(content_frame, text="", font=("Segoe UI", 12))
status_label.pack(pady=6)

# ----- BEÁLLÍTÁSOK TAB (scrollable) -----
settings_tab = tb.Frame(notebook)
notebook.add(settings_tab, text="Beállítások")

# Canvas + Scrollbar for settings tab
settings_canvas = tb.Canvas(settings_tab)
settings_canvas.pack(side="left", fill="both", expand=True)
settings_scrollbar = tb.Scrollbar(
    settings_tab, orient="vertical", command=settings_canvas.yview
)
settings_scrollbar.pack(side="right", fill="y")
settings_canvas.configure(yscrollcommand=settings_scrollbar.set)

set_frame = tb.Frame(settings_canvas, padding=30)
set_frame_id = settings_canvas.create_window((0, 0), window=set_frame, anchor="nw")


def _on_settings_frame_configure(event):
    settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))


set_frame.bind("<Configure>", _on_settings_frame_configure)


def _on_settings_mousewheel(event):
    settings_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


settings_canvas.bind_all("<MouseWheel>", _on_settings_mousewheel)

# Megjelenés
ui_group = tb.Labelframe(set_frame, text="🎨 Megjelenés", padding=12)
ui_group.pack(fill=X, pady=8)

tb.Label(ui_group, text="Téma:").grid(row=0, column=0, sticky=W, padx=4, pady=4)
themes = [
    "minty",
    "flatly",
    "pulse",
    "superhero",
    "cosmo",
    "cyborg",
    "journal",
    "litera",
    "lumen",
    "materia",
    "sandstone",
    "simplex",
    "solar",
    "united",
    "yeti",
    "morph",
]
theme_var = tk.StringVar(value=SETTINGS.get("theme", "minty"))
tb.Combobox(
    ui_group, textvariable=theme_var, values=themes, state="readonly", width=18
).grid(row=0, column=1, sticky=W, padx=4, pady=4)
tb.Button(
    ui_group,
    text="Alkalmaz",
    bootstyle=PRIMARY,
    command=lambda: apply_theme(theme_var.get()),
).grid(row=0, column=2, padx=6, pady=4)

tb.Label(ui_group, text="UI skála (%):").grid(row=1, column=0, sticky=W, padx=4, pady=4)
scale_var = tk.IntVar(value=int(SETTINGS.get("ui_scale_pct", 100)))
tb.Spinbox(
    ui_group, from_=50, to=200, increment=25, textvariable=scale_var, width=10
).grid(row=1, column=1, sticky=W, padx=4, pady=4)
tb.Button(
    ui_group,
    text="Alkalmaz",
    bootstyle=PRIMARY,
    command=lambda: apply_scale(scale_var.get()),
).grid(row=1, column=2, padx=6, pady=4)

tb.Label(ui_group, text="Nyelv:").grid(row=2, column=0, sticky=W, padx=4, pady=4)
lang_var = tk.StringVar(value=SETTINGS.get("language", "hu"))
tb.Combobox(
    ui_group, textvariable=lang_var, values=["hu", "en"], state="readonly", width=10
).grid(row=2, column=1, sticky=W, padx=4, pady=4)

# Mappák
path_group = tb.Labelframe(set_frame, text="📂 Mappák", padding=12)
path_group.pack(fill=X, pady=8)

tb.Label(path_group, text="TS fájlok mappa:").grid(
    row=0, column=0, sticky=W, padx=4, pady=4
)
ts_folder_var = tk.StringVar(value=SETTINGS.get("ts_folder", str(Path.cwd())))
ts_entry = tb.Entry(path_group, textvariable=ts_folder_var, width=60)
ts_entry.grid(row=0, column=1, sticky=W, padx=4, pady=4)
tb.Button(path_group, text="Választ…", command=lambda: choose_dir(ts_entry)).grid(
    row=0, column=2, padx=6, pady=4
)

tb.Label(path_group, text="Kimeneti mappa:").grid(
    row=1, column=0, sticky=W, padx=4, pady=4
)
output_folder_var = tk.StringVar(value=SETTINGS.get("output_folder", ""))
out_entry = tb.Entry(path_group, textvariable=output_folder_var, width=60)
out_entry.grid(row=1, column=1, sticky=W, padx=4, pady=4)
tb.Button(path_group, text="Választ…", command=lambda: choose_dir(out_entry)).grid(
    row=1, column=2, padx=6, pady=4
)

tb.Label(path_group, text="Backup mappa:").grid(
    row=2, column=0, sticky=W, padx=4, pady=4
)
backup_folder_var = tk.StringVar(value=SETTINGS.get("backup_folder", ""))
bak_entry = tb.Entry(path_group, textvariable=backup_folder_var, width=60)
bak_entry.grid(row=2, column=1, sticky=W, padx=4, pady=4)
tb.Button(path_group, text="Választ…", command=lambda: choose_dir(bak_entry)).grid(
    row=2, column=2, padx=6, pady=4
)

backup_enabled_var = tk.BooleanVar(value=bool(SETTINGS.get("backup_enabled", False)))
tb.Checkbutton(
    path_group,
    text="Backup készítése mentéskor",
    variable=backup_enabled_var,
    bootstyle="round-toggle",
).grid(row=3, column=1, sticky=W, padx=4, pady=4)

# Ügyfélkódok – ÚJ LISTAKEZELŐ UI
codes_group = tb.Labelframe(
    set_frame, text="👥 Ügyfélkódok (alapértelmezett lista)", padding=12
)
codes_group.pack(fill=X, pady=8)

all_codes_available = _load_all_client_codes_sorted()
initial_defaults = _initial_default_codes_for_settings_ui(all_codes_available)

# bal: listbox a kiválasztott alapértelmezett kódokkal
default_codes_listbox = tk.Listbox(
    codes_group, selectmode="extended", activestyle="dotbox", height=8
)
default_codes_listbox.grid(
    row=0, column=0, rowspan=6, sticky="nsew", padx=(4, 4), pady=4
)
codes_group.grid_columnconfigure(0, weight=1)
codes_group.grid_rowconfigure(0, weight=1)


# show the default codes alphabetically (accent-insensitive)
for code in sorted(initial_defaults, key=remove_accents):
    default_codes_listbox.insert(tk.END, code)

# jobb: műveleti gombok
btn_col = tb.Frame(codes_group)
btn_col.grid(row=0, column=1, sticky="ns", padx=(6, 6), pady=4)


def _lb_move(delta: int):
    sel = list(default_codes_listbox.curselection())
    if not sel:
        return
    # több elem mozgatása sorrendhelyesen
    items = list(default_codes_listbox.get(0, tk.END))
    new_sel = []
    rng = range(len(items))
    order = sel if delta < 0 else sel[::-1]
    for i in order:
        j = i + delta
        if j < 0 or j >= len(items):
            continue
        items[i], items[j] = items[j], items[i]
        new_sel.append(j)
    default_codes_listbox.delete(0, tk.END)
    for x in items:
        default_codes_listbox.insert(tk.END, x)
    default_codes_listbox.selection_clear(0, tk.END)
    for j in new_sel[::-1]:
        default_codes_listbox.selection_set(j)


tb.Button(
    btn_col, text="Fel ▲", bootstyle=SECONDARY, command=lambda: _lb_move(-1)
).pack(fill=X, pady=2)
tb.Button(btn_col, text="Le ▼", bootstyle=SECONDARY, command=lambda: _lb_move(+1)).pack(
    fill=X, pady=2
)


def _lb_remove():
    sel = list(default_codes_listbox.curselection())
    for i in reversed(sel):
        default_codes_listbox.delete(i)


tb.Button(btn_col, text="Töröl", bootstyle=DANGER, command=_lb_remove).pack(
    fill=X, pady=(8, 2)
)


def _lb_clear():
    default_codes_listbox.delete(0, tk.END)


tb.Button(btn_col, text="Ürít", bootstyle="outline-danger", command=_lb_clear).pack(
    fill=X, pady=2
)

# alul: Hozzáadás sor – combobox (szabad gépelés engedett)
add_row = tb.Frame(codes_group)
add_row.grid(row=6, column=0, columnspan=2, sticky="ew", padx=4, pady=(8, 4))
tb.Label(add_row, text="Új kód:").pack(side=LEFT)
add_var = tk.StringVar()
add_combo = tb.Combobox(
    add_row, textvariable=add_var, values=all_codes_available, width=30
)
add_combo.pack(side=LEFT, padx=6)


def _lb_add():
    code = (add_var.get() or "").strip()
    if not code:
        return
    # ha van master lista, ellenőrizzük – ha nincs, engedjük a szabad beírást
    if all_codes_available and code not in all_codes_available:
        messagebox.showwarning("Ügyfélkód", f"Ismeretlen kód: {code}")
        return
    existing = list(default_codes_listbox.get(0, tk.END))
    if code in existing:
        # fókuszáljuk meg
        idx = existing.index(code)
        default_codes_listbox.selection_clear(0, tk.END)
        default_codes_listbox.selection_set(idx)
        default_codes_listbox.see(idx)
        return
    # insert preserving alphabetical order (accent-insensitive)
    insert_idx = 0
    for i, item in enumerate(existing):
        if remove_accents(code) > remove_accents(item):
            insert_idx = i + 1
        else:
            break
    default_codes_listbox.insert(insert_idx, code)
    default_codes_listbox.selection_clear(0, tk.END)
    default_codes_listbox.selection_set(insert_idx)
    add_var.set("")


tb.Button(add_row, text="Hozzáad", bootstyle=PRIMARY, command=_lb_add).pack(
    side=LEFT, padx=6
)

# “Utolsó választás megjegyzése” kapcsoló
remember_last_var = tk.BooleanVar(
    value=bool(SETTINGS.get("remember_last_selection", True))
)
tb.Checkbutton(
    codes_group,
    text="Utolsó választás megjegyzése",
    variable=remember_last_var,
    bootstyle="round-toggle",
).grid(row=7, column=0, columnspan=2, sticky=W, padx=4, pady=(6, 0))

# Folyamat / értesítések
run_group = tb.Labelframe(set_frame, text="⚙️ Folyamat és értesítések", padding=12)
run_group.pack(fill=X, pady=8)

auto_open_output_var = tk.BooleanVar(
    value=bool(SETTINGS.get("auto_open_output_on_success", False))
)
tb.Checkbutton(
    run_group,
    text="Kimeneti fájl automatikus megnyitása siker esetén",
    variable=auto_open_output_var,
    bootstyle="round-toggle",
).grid(row=0, column=0, columnspan=2, sticky=W, padx=4, pady=4)

auto_open_details_var = tk.BooleanVar(
    value=bool(SETTINGS.get("auto_open_details_on_error", True))
)
tb.Checkbutton(
    run_group,
    text="Részletek automatikus megnyitása hiba esetén",
    variable=auto_open_details_var,
    bootstyle="round-toggle",
).grid(row=1, column=0, columnspan=2, sticky=W, padx=4, pady=4)

tb.Label(run_group, text="Popup automatikus bezárás (mp):").grid(
    row=2, column=0, sticky=W, padx=4, pady=4
)
popup_autoclose_var = tk.StringVar(
    value=str(int(SETTINGS.get("popup_autoclose_sec", 0)))
)
tb.Entry(run_group, textvariable=popup_autoclose_var, width=8).grid(
    row=2, column=1, sticky=W, padx=4, pady=4
)

sound_var = tk.BooleanVar(value=bool(SETTINGS.get("sound_enabled", True)))
tb.Checkbutton(
    run_group,
    text="Hangjelzés művelet végén",
    variable=sound_var,
    bootstyle="round-toggle",
).grid(row=3, column=0, columnspan=2, sticky=W, padx=4, pady=4)

# --- Napi emlékeztető
daily_reminder_enabled_var = tk.BooleanVar(
    value=bool(SETTINGS.get("daily_reminder_enabled", False))
)
tb.Checkbutton(
    run_group,
    text="Napi emlékeztető engedélyezése",
    variable=daily_reminder_enabled_var,
    bootstyle="round-toggle",
).grid(row=4, column=0, columnspan=2, sticky=W, padx=4, pady=4)

tb.Label(run_group, text="Emlékeztető időpont (HH:MM):").grid(
    row=5, column=0, sticky=W, padx=4, pady=4
)
daily_reminder_time_var = tk.StringVar(
    value=str(SETTINGS.get("daily_reminder_time", "18:00"))
)
tb.Entry(run_group, textvariable=daily_reminder_time_var, width=10).grid(
    row=5, column=1, sticky=W, padx=4, pady=4
)

# (Automatikus riport UI továbbra is kikommentelve – jelen állapot megőrzése)

# Gombok
btn_row = tb.Frame(set_frame)
btn_row.pack(fill=X, pady=(10, 0))
tb.Button(
    btn_row, text="Beállítások mentése", bootstyle=SUCCESS, command=settings_save
).pack(side=RIGHT)

# Exit
exit_frame = tb.Frame(app)
exit_frame.pack(side=BOTTOM, fill=X, pady=10)
tb.Button(exit_frame, text="Kilépés", bootstyle=DANGER, command=app.destroy).pack(
    pady=6, ipadx=10, ipady=6
)

# Start UI pump + időzítők
app.after(100, pump_ui)
app.after(150, refresh_dashboard)


# napi emlékeztető (placeholder)
def reminder_tick():
    try:
        if SETTINGS.get("daily_reminder_enabled"):
            _ = parse_hhmm(str(SETTINGS.get("daily_reminder_time", "18:00")))
            # hely itt, ha később szeretnél konkrét értesítést
    except Exception:
        pass
    finally:
        app.after(60_000, reminder_tick)


app.after(1000, reminder_tick)
app.after(2000, weekly_report_tick)  # ütemezett heti riport

app.mainloop()
