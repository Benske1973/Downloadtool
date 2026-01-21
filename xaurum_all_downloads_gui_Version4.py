# xaurum_all_downloads_gui.py
# ---------------------------------------------------------------
# Xaurum downloader GUI + SILENT runner met slimme MFA-check.
# - Bij opstart: smart_bootstrap() controleert of MFA/login nodig is.
#   * Nodig?  Edge zichtbaar, gebruiker rondt MFA af; daarna X_HEADLESS=1.
#   * Niet nodig/al ingelogd? Direct X_HEADLESS=1 en downloads starten automatisch.
# - Na downloads: xaurum_converter.py wordt automatisch uitgevoerd.
# - Silent modus doet vooraf dezelfde auth-check.
#
# Vereist: smart_auth_bootstrap.py (zelfde folder) met functie smart_bootstrap().
# ---------------------------------------------------------------

from __future__ import annotations
import sys
from pathlib import Path
HERE = Path(__file__).parent.resolve()               # ...\XaurumTools
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
if str(HERE / "Scripts") not in sys.path:
    sys.path.insert(0, str(HERE / "Scripts"))
import io
import os
import time
import runpy
import threading
import queue
import argparse
import subprocess
from pathlib import Path
from datetime import datetime

# >>> MFA bootstrap (moet aanwezig zijn)
from smart_auth_bootstrap import smart_bootstrap

# ================ Config ================
SCRIPT_NAMES = [
    "xaurum_dispatcher_certificates.py",
    "xaurum_dispatcher_formations.py",
    "xaurum_dispatcher_certified_report.py",
    "xaurum_download_competences.py",
]

# Converter script (in zelfde map als dit GUI script)
CONVERTER_SCRIPT = "xaurum_converter.py"

RETRY_POLICY_TEXT = (
    "Elke downloader probeert max. 3x (1x + 2 herpogingen)."
    "Bij definitieve mislukking wordt een e-mail verstuurd."
)

# --------- Config ophalen uit xaurum_common (downloadmap + auth state) -----------
def _get_config_from_common(scripts_dir: Path | None) -> tuple[Path, Path]: 
    """
    Haal DL_DIR en AUTH_STATE op uit xaurum_common; val anders terug op defaults.
    """
    default_dl = Path.cwd() / "downloads_xaurum"
    default_auth = Path(os.environ.get("APPDATA", str(Path.home()))) / "XaurumUploader" / "xaurum_auth_state.json"

    if not scripts_dir: 
        return default_dl, default_auth
    try:
        sys.path.insert(0, str(scripts_dir))
        from xaurum_common import DL_DIR, AUTH_STATE  # type: ignore
        return DL_DIR, AUTH_STATE
    except Exception:
        return default_dl, default_auth

# ================ Silent runner ================
def run_scripts_silent(base_dir: Path, scripts_dir: Path, log_dir: Path) -> int:
    """
    Run zonder GUI. Doet eerst de slimme login/MFA-check.
    """
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"launcher_{datetime.now():%Y%m%d_%H%M%S}.log"

    def w(line:  str):
        print(line, flush=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")

    w("============================================================")
    w("Xaurum launcher (SILENT)")
    w("============================================================")
    w(f"Base map :  {base_dir}")
    w(f"Scripts  : {scripts_dir}")
    w(f"Log file : {log_path}")
    w(f"Datum    : {datetime.now():%Y-%m-%d %H:%M:%S}")
    w("------------------------------------------------------------")
    w(f"Retry policy:  {RETRY_POLICY_TEXT}")
    w("------------------------------------------------------------")

    # 0) Slimme auth-check (opent Edge zichtbaar indien MFA nodig is)
    w("[auth] Start slimme login-check…")
    ok = False
    try: 
        ok = smart_bootstrap()
    except Exception as e:
        w(f"[auth] Fout tijdens auth-check: {e!r}")

    if not ok: 
        w("❌ Login/MFA niet afgerond. Stop.")
        return 1

    # 1) afdwingen headless uitvoering in child scripts
    os.environ["X_HEADLESS"] = "1"
    # 2) Zorg dat child scripts xaurum_common opnieuw met env oppikken
    sys.path.insert(0, str(scripts_dir))

    results:  list[tuple[str, int, float]] = []

    for i, script in enumerate(SCRIPT_NAMES, start=1):
        start = time.time()
        w("")
        w("------------------------------------------------------------")
        w(f"▶ [{i}/{len(SCRIPT_NAMES)}] Start {script}")
        w("------------------------------------------------------------")

        rc = 1
        old_argv, old_out, old_err = sys.argv[: ], sys.stdout, sys.stderr
        try:
            # route stdout/err naar log
            class Tee(io.TextIOBase):
                def write(self, s):
                    if s:
                        for ln in s.splitlines():
                            w(ln)
                def flush(self):  # noqa: D401
                    return None

            sys.stdout = Tee()
            sys.stderr = Tee()
            sys.argv = [str(scripts_dir / script)]
            runpy.run_path(str(scripts_dir / script), run_name="__main__")
            rc = 0
        except SystemExit as e: 
            try:
                rc = int(e.code) if e.code is not None else 0
            except Exception:
                rc = 1
        except Exception as e: 
            w(f"💥 Exception in {script}: {e}")
            rc = 1
        finally: 
            sys.argv = old_argv
            sys.stdout = old_out
            sys.stderr = old_err

        dt = time.time() - start
        if rc == 0:
            w(f"✅ Klaar (code=0): {script}  — duur: {dt:.1f}s")
        else:
            w(f"❌ Klaar (code={rc}): {script}  — duur: {dt:.1f}s")
        results.append((script, rc, dt))

        # kleine pauze tussen scripts
        time.sleep(1.0)

    # ========== CONVERTER ==========
    w("")
    w("============================================================")
    w("▶ CONVERTER starten")
    w("============================================================")
    converter_path = base_dir / CONVERTER_SCRIPT
    converter_rc = 1
    if converter_path.exists():
        start = time.time()
        old_argv, old_out, old_err = sys.argv[: ], sys.stdout, sys.stderr
        try:
            class Tee(io.TextIOBase):
                def write(self, s):
                    if s: 
                        for ln in s.splitlines():
                            w(ln)
                def flush(self):
                    return None

            sys.stdout = Tee()
            sys.stderr = Tee()
            sys.argv = [str(converter_path)]
            runpy.run_path(str(converter_path), run_name="__main__")
            converter_rc = 0
        except SystemExit as e:
            try:
                converter_rc = int(e.code) if e.code is not None else 0
            except Exception:
                converter_rc = 1
        except Exception as e:
            w(f"💥 Exception in converter: {e}")
            converter_rc = 1
        finally:
            sys.argv = old_argv
            sys.stdout = old_out
            sys.stderr = old_err

        dt = time.time() - start
        if converter_rc == 0:
            w(f"✅ Converter klaar (code=0) — duur: {dt:.1f}s")
        else:
            w(f"❌ Converter fout (code={converter_rc}) — duur: {dt:.1f}s")
    else:
        w(f"⚠️ Converter niet gevonden:  {converter_path}")

    # Samenvatting
    w("")
    w("============================================================")
    w("EINDOVERZICHT")
    w("============================================================")
    okc = 0
    for s, rc, _ in results:
        if rc == 0:
            w(f"OK   {s}")
            okc += 1
        else:
            w(f"ERR  {s}")
    if converter_rc == 0:
        w(f"OK   {CONVERTER_SCRIPT}")
    else:
        w(f"ERR  {CONVERTER_SCRIPT}")
    w("------------------------------------------------------------")
    total_ok = okc + (1 if converter_rc == 0 else 0)
    total_scripts = len(SCRIPT_NAMES) + 1
    w(f"Gereed. {total_ok}/{total_scripts} succesvol.")
    w(f"Logbestand: {log_path}")

    # exit code:  0 als alles ok, anders 1
    return 0 if total_ok == total_scripts else 1

# ================ GUI ================
import tkinter as tk
from tkinter import ttk, messagebox


class StreamToQueue(io.TextIOBase):
    def __init__(self, q: queue.Queue[str], prefix: str = ""):
        self.q, self.prefix = q, prefix

    def write(self, s):
        if not s:
            return
        for line in s.splitlines():
            self.q.put(self.prefix + line)

    def flush(self):
        return None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Xaurum Downloads + Converter")
        self.geometry("900x680")
        self.resizable(False, False)

        self.base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        candidates = [self.base / "Scripts", self.base.parent / "Scripts", Path.cwd() / "Scripts"]
        self.scripts_dir = next((p for p in candidates if p.exists() and p.is_dir()), None)

        # download map en auth state
        self.download_dir, self.auth_state = _get_config_from_common(self.scripts_dir)
        self.log_q:  queue.Queue[str] = queue.Queue()
        self.running = False
        self.cancel = False

        # Alle taken:  4 downloaders + 1 converter
        self.all_tasks = SCRIPT_NAMES + [CONVERTER_SCRIPT]

        # Titel + retry-policy
        ttk.Label(self, text="Automatische Xaurum downloads + conversie", font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
        ttk.Label(self, text=f"Retry policy: {RETRY_POLICY_TEXT}", foreground="#555").pack(pady=(0, 4))

        # Knoppen frame (zonder Start-knop)
        top = ttk.Frame(self)
        top.pack(fill="x", padx=12)
        self.btn_cancel = ttk.Button(top, text="⏹ Annuleer", width=14, state="disabled", command=self.ask_cancel)
        self.btn_cancel.pack(side="left", padx=(0, 8))
        self.btn_retry = ttk.Button(top, text="🔐 Probeer login opnieuw", width=24,
                                    command=self.start_auth_check, state="disabled")
        self.btn_retry.pack(side="left", padx=(0, 8))

        # pad + openen
        path_frame = ttk.Frame(self)
        path_frame.pack(fill="x", padx=12, pady=(6, 6))
        ttk.Label(path_frame, text=f"Downloadmap: {self.download_dir}", anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(path_frame, text="Openen", command=self.open_folder).pack(side="right")

        # status + progress
        self.status = ttk.Label(self, anchor="w", text="Controleren of login/MFA vereist is…")
        self.status.pack(fill="x", padx=12, pady=(0, 2))
        self.pb = ttk.Progressbar(self, maximum=len(self.all_tasks), mode="indeterminate")
        self.pb.pack(fill="x", padx=12)
        self.pb.start(12)

        # lijst met alle taken (downloaders + converter)
        self.tree = ttk.Treeview(self, show="tree")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        self.nodes = []
        for s in SCRIPT_NAMES: 
            self.nodes.append(self.tree.insert("", "end", text=f"📥 {s}  —  wacht…", open=True))
        # Converter apart markeren
        self.nodes.append(self.tree.insert("", "end", text=f"🔄 {CONVERTER_SCRIPT}  —  wacht…", open=True))

        ttk.Label(self, text="Log:", anchor="w").pack(fill="x", padx=12)
        self.txt = tk.Text(self, wrap="word", height=12)
        self.txt.pack(fill="both", expand=False, padx=12, pady=(0, 12))

        self.after(120, self.drain_log)
        self.preflight()

        # Start de auth-check (en daarna automatisch downloaden)
        self.after(200, self.start_auth_check)

    # ---------- kleine helpers ----------
    def log(self, msg:  str):
        self.log_q.put(msg)

    def drain_log(self):
        try:
            while True:
                line = self.log_q.get_nowait()
                self.txt.insert("end", line + "\n")
                self.txt.see("end")
        except queue.Empty:
            pass
        self.after(120, self.drain_log)

    # ---------- UI acties ----------
    def open_folder(self):
        target = str(self.download_dir)
        try:
            if os.name == "nt":
                subprocess.Popen(["explorer", target])
            else:
                subprocess.Popen(["xdg-open", target])
        except Exception as e:
            messagebox.showerror("Fout", f"Kon map niet openen:\n{target}\n\n{e}")

    def preflight(self):
        if not self.scripts_dir:
            messagebox.showerror("Fout", "Scripts-map niet gevonden.\nPlaats deze launcher naast 'Scripts'.")
            self.destroy()
            return
        miss = [s for s in SCRIPT_NAMES if not (self.scripts_dir / s).exists()]
        if miss:
            messagebox.showerror("Fout", "Ontbrekende scripts:\n- " + "\n- ".join(miss))
        # Check converter
        converter_path = self.base / CONVERTER_SCRIPT
        if not converter_path.exists():
            messagebox.showwarning("Waarschuwing", f"Converter niet gevonden:\n{converter_path}\n\nDownloads werken, maar conversie wordt overgeslagen.")

    # ---------- Auth-check flow ----------
    def start_auth_check(self):
        """Run de slimme login-check in een background thread."""
        self.btn_retry.config(state="disabled")
        self.btn_cancel.config(state="disabled")
        self.pb.configure(mode="indeterminate")
        self.pb.start(12)
        self.status.config(text="Bezig met login-check… (browser kan zichtbaar openen)")

        t = threading.Thread(target=self._auth_check_worker, daemon=True)
        t.start()

    def _auth_check_worker(self):
        ok = False
        try:
            ok = smart_bootstrap()  # toont Edge zichtbaar indien MFA nodig is; zet X_HEADLESS=1 nadien
        except Exception as e:
            self.after(0, lambda: self.log(f"ERR auth:  {e!r}"))

        if ok:
            def _ok_ui():
                self.pb.stop()
                self.pb.configure(mode="determinate")
                self.pb["value"] = 0
                self.status.config(text="Login OK. Downloads starten automatisch...")
                self.btn_retry.config(state="disabled")
                self.log("✅ Login geslaagd, downloads starten automatisch...")
                # Start automatisch met downloaden! 
                self.after(500, self.start_downloads)
            self.after(0, _ok_ui)
        else:
            def _fail_ui():
                self.pb.stop()
                self.pb.configure(mode="determinate")
                self.pb["value"] = 0
                self.status.config(text="Login niet afgerond.Klik 'Probeer login opnieuw'.")
                self.btn_retry.config(state="normal")
                self.log("⚠️ Login/MFA niet afgerond of time-out.Probeer opnieuw. "
                         "Laat Edge open tot Xaurum volledig geladen is.")
            self.after(0, _fail_ui)

    # ---------- Downloader flow ----------
    def start_downloads(self):
        """Start de downloads automatisch na succesvolle login."""
        if self.running or not self.scripts_dir:
            return
        self.running = True
        self.cancel = False
        self.btn_cancel.config(state="normal")
        self.btn_retry.config(state="disabled")
        self.status.config(text="Bezig met downloaden…")
        self.pb.configure(mode="determinate")
        self.pb["value"] = 0
        for i, n in enumerate(self.nodes):
            if i < len(SCRIPT_NAMES):
                self.tree.item(n, text=f"📥 {self.all_tasks[i]}  —  wacht…")
            else:
                self.tree.item(n, text=f"🔄 {self.all_tasks[i]}  —  wacht…")
        threading.Thread(target=self.worker, daemon=True).start()

    def ask_cancel(self):
        if messagebox.askyesno("Annuleer", "Wil je de reeks stoppen?"):
            self.cancel = True
            self.btn_cancel.config(state="disabled")
            self.status.config(text="Geannuleerd…")

    def run_script(self, path: Path) -> int:
        # xaurum_common opnieuw laten laden met env (X_HEADLESS / profiel)
        if "xaurum_common" in sys.modules:
            del sys.modules["xaurum_common"]
        sys.path.insert(0, str(self.scripts_dir))

        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = StreamToQueue(self.log_q)
        sys.stderr = StreamToQueue(self.log_q, "ERR:  ")
        old_argv = sys.argv[:]
        sys.argv = [str(path)]
        try:
            runpy.run_path(str(path), run_name="__main__")
            return 0
        except SystemExit as e:
            try:
                return int(e.code) if e.code is not None else 0
            except Exception: 
                return 1
        except Exception as e: 
            self.log(f"💥 Exception in {path.name}: {e}")
            return 1
        finally: 
            sys.argv = old_argv
            sys.stdout = old_out
            sys.stderr = old_err

    def worker(self):
        log_dir = self.download_dir if self.download_dir else (self.base / "downloads_xaurum")
        log_dir.mkdir(parents=True, exist_ok=True)
        logfile = (log_dir / f"launcher_{datetime.now():%Y%m%d_%H%M%S}.log").open("w", encoding="utf-8")

        def w(line: str):
            self.log(line)
            logfile.write(line + "\n")
            logfile.flush()

        w(f"📂 Downloadmap: {log_dir}")
        w(f"🔧 X_HEADLESS: {os.environ.get('X_HEADLESS', 'NIET GEZET!')}")
        w(f"Retry policy: {RETRY_POLICY_TEXT}")

        results:  list[tuple[str, int]] = []
        
        # ========== DOWNLOADERS ==========
        for idx, name in enumerate(SCRIPT_NAMES):
            if self.cancel:
                break
            path = self.scripts_dir / name
            self.tree.item(self.nodes[idx], text=f"📥 {name}  —  bezig…")
            w(f"\n{'='*64}\n▶ [{idx+1}/{len(SCRIPT_NAMES)}] {name}\n{'='*64}")
            t0 = time.time()
            rc = self.run_script(path)
            dt = time.time() - t0
            if rc == 0:
                self.tree.item(self.nodes[idx], text=f"📥 {name}  —  �� OK")
                w(f"✅ Klaar in {dt:.1f}s")
            else:
                self.tree.item(self.nodes[idx], text=f"📥 {name}  —  ❌ ERROR ({rc})")
                w(f"❌ Fout (code {rc}) na {dt:.1f}s")
            self.pb["value"] = idx + 1
            results.append((name, rc))
            time.sleep(0.5)

        # ========== CONVERTER ==========
        converter_idx = len(SCRIPT_NAMES)
        converter_rc = 1
        if not self.cancel:
            converter_path = self.base / CONVERTER_SCRIPT
            self.tree.item(self.nodes[converter_idx], text=f"🔄 {CONVERTER_SCRIPT}  —  bezig…")
            self.status.config(text="Bezig met conversie…")
            w(f"\n{'='*64}\n▶ CONVERTER:  {CONVERTER_SCRIPT}\n{'='*64}")
            
            if converter_path.exists():
                t0 = time.time()
                converter_rc = self.run_script(converter_path)
                dt = time.time() - t0
                if converter_rc == 0:
                    self.tree.item(self.nodes[converter_idx], text=f"🔄 {CONVERTER_SCRIPT}  —  ✅ OK")
                    w(f"✅ Converter klaar in {dt:.1f}s")
                else: 
                    self.tree.item(self.nodes[converter_idx], text=f"🔄 {CONVERTER_SCRIPT}  —  ❌ ERROR ({converter_rc})")
                    w(f"❌ Converter fout (code {converter_rc}) na {dt:.1f}s")
            else:
                self.tree.item(self.nodes[converter_idx], text=f"🔄 {CONVERTER_SCRIPT}  —  ⚠️ NIET GEVONDEN")
                w(f"⚠️ Converter niet gevonden: {converter_path}")
            
            self.pb["value"] = converter_idx + 1
            results.append((CONVERTER_SCRIPT, converter_rc))

        # ========== SAMENVATTING ==========
        w("\n============================================================")
        w("EINDOVERZICHT")
        w("============================================================")
        okc = 0
        for s, rc in results:
            if rc == 0:
                w(f"OK   {s}")
                okc += 1
            else: 
                w(f"ERR  {s}")
        w("------------------------------------------------------------")
        w(f"Gereed.{okc}/{len(results)} succesvol.")
        logfile.close()

        self.btn_cancel.config(state="disabled")
        self.running = False
        if self.cancel:
            self.status.config(text="Geannuleerd.")
        else: 
            self.status.config(text=f"Klaar! {okc}/{len(results)} succesvol.")

# ================ Entry ================
def main():
    # Bepaal base & Scripts
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    scripts_dir = next((p for p in [base / "Scripts", base.parent / "Scripts", Path.cwd() / "Scripts"] if p.exists() and p.is_dir()), None)
    dl_dir, _ = _get_config_from_common(scripts_dir)

    # CLI args
    ap = argparse.ArgumentParser(description="Xaurum downloader + converter (GUI of silent).")
    ap.add_argument("--silent", action="store_true", help="Run zonder GUI (voor Taakplanner).")
    ap.add_argument("--logdir", type=str, default=str(dl_dir), help="Map voor logbestanden (default = downloadmap).")
    args = ap.parse_args()

    if args.silent:
        rc = run_scripts_silent(base, scripts_dir or (base / "Scripts"), Path(args.logdir))
        raise SystemExit(rc)
    else:
        # GUI
        import tkinter  # lazy import om headless errors te vermijden bij --silent
        App().mainloop()


if __name__ == "__main__": 
    main()