from __future__ import annotations

import json
import os
from pathlib import Path
import queue
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name.casefold() == "scripts" else SCRIPT_DIR
SCRIPTS_DIR = PROJECT_ROOT / "Scripts"
WORKING_FILES_DIR = PROJECT_ROOT / "Working Files"
OUTPUT_WORKBOOK = PROJECT_ROOT / "Consolidated New Biz Report.xlsx"
VENV_PYTHON = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"
REQUIREMENTS = SCRIPTS_DIR / "requirements.txt"
CONFIG_DIR = PROJECT_ROOT / "config"
EPIC_CREDENTIALS_PATH = CONFIG_DIR / "epic_credentials.json"
BIGNITION_CREDENTIALS_PATH = CONFIG_DIR / "bignition_credentials.json"

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

PIPELINES: dict[str, list[str]] = {
    "Full Pipeline (Bignition + EPIC + All Tabs)": [
        "main.py",
        "epic_report.py",
        "data_consolidation.py",
        "new_biz_tabs.py",
        "written_business_ytd.py",
    ],
    "Bignition + Consolidation + Tabs": [
        "main.py",
        "data_consolidation.py",
        "new_biz_tabs.py",
        "written_business_ytd.py",
    ],
    "EPIC + Consolidation + Tabs": [
        "epic_report.py",
        "data_consolidation.py",
        "new_biz_tabs.py",
        "written_business_ytd.py",
    ],
    "Consolidation Only": [
        "data_consolidation.py",
    ],
    "New Biz Tabs Only": [
        "new_biz_tabs.py",
    ],
    "Written Business Only": [
        "written_business_ytd.py",
    ],
}


def timestamp() -> str:
    return datetime.now().strftime("%H:%M:%S")


class AutomationUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("New Biz Report Automation")
        self.root.geometry("1100x700")
        self.root.minsize(1000, 600)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.current_process: subprocess.Popen[str] | None = None
        self.stop_requested = False

        self.pipeline_var = tk.StringVar(value=list(PIPELINES.keys())[0])
        self.setup_before_run_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Idle")
        self.epic_usercode_var = tk.StringVar(value="")
        self.epic_password_var = tk.StringVar(value="")
        self.bignition_username_var = tk.StringVar(value="")
        self.bignition_password_var = tk.StringVar(value="")
        self.show_epic_password_var = tk.BooleanVar(value=False)
        self.show_bignition_password_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.load_saved_credentials()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._drain_log_queue)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(container)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Pipeline:").pack(side=tk.LEFT)
        self.pipeline_combo = ttk.Combobox(
            top,
            textvariable=self.pipeline_var,
            values=list(PIPELINES.keys()),
            state="readonly",
            width=48,
        )
        self.pipeline_combo.pack(side=tk.LEFT, padx=(8, 12))

        self.run_btn = ttk.Button(top, text="Run Selected", command=self.start_pipeline)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(top, text="Stop", command=self.stop_pipeline, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.setup_btn = ttk.Button(top, text="Setup Environment", command=self.start_setup_only)
        self.setup_btn.pack(side=tk.LEFT)

        options = ttk.Frame(container)
        options.pack(fill=tk.X, pady=(10, 10))
        ttk.Checkbutton(
            options,
            text="Run setup before pipeline (venv, pip install, playwright install)",
            variable=self.setup_before_run_var,
        ).pack(side=tk.LEFT)

        creds = ttk.LabelFrame(container, text="EPIC Credentials (Auto Login)", padding=10)
        creds.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(creds, text="Usercode").grid(row=0, column=0, sticky="w")
        ttk.Entry(creds, textvariable=self.epic_usercode_var, width=28).grid(row=0, column=1, sticky="w", padx=(8, 14))
        ttk.Label(creds, text="Password").grid(row=0, column=2, sticky="w")
        self.epic_password_entry = ttk.Entry(creds, textvariable=self.epic_password_var, show="*", width=28)
        self.epic_password_entry.grid(row=0, column=3, sticky="w", padx=(8, 8))
        ttk.Checkbutton(
            creds,
            text="Show Password",
            variable=self.show_epic_password_var,
            command=self.toggle_epic_password_visibility,
        ).grid(row=0, column=4, sticky="w", padx=(0, 12))
        ttk.Button(creds, text="Save Credentials", command=self.save_credentials).grid(row=0, column=5, sticky="w", padx=(0, 8))
        ttk.Button(creds, text="Clear", command=self.clear_credentials).grid(row=0, column=6, sticky="w")
        ttk.Label(
            creds,
            text=f"Stored locally: {EPIC_CREDENTIALS_PATH}",
        ).grid(row=1, column=0, columnspan=7, sticky="w", pady=(8, 0))

        bignition_creds = ttk.LabelFrame(container, text="Bignition Credentials (Auto Login)", padding=10)
        bignition_creds.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(bignition_creds, text="Username").grid(row=0, column=0, sticky="w")
        ttk.Entry(bignition_creds, textvariable=self.bignition_username_var, width=28).grid(
            row=0, column=1, sticky="w", padx=(8, 14)
        )
        ttk.Label(bignition_creds, text="Password").grid(row=0, column=2, sticky="w")
        self.bignition_password_entry = ttk.Entry(bignition_creds, textvariable=self.bignition_password_var, show="*", width=28)
        self.bignition_password_entry.grid(row=0, column=3, sticky="w", padx=(8, 8))
        ttk.Checkbutton(
            bignition_creds,
            text="Show Password",
            variable=self.show_bignition_password_var,
            command=self.toggle_bignition_password_visibility,
        ).grid(row=0, column=4, sticky="w", padx=(0, 12))
        ttk.Button(bignition_creds, text="Save Credentials", command=self.save_bignition_credentials).grid(
            row=0, column=5, sticky="w", padx=(0, 8)
        )
        ttk.Button(bignition_creds, text="Clear", command=self.clear_bignition_credentials).grid(row=0, column=6, sticky="w")
        ttk.Label(
            bignition_creds,
            text=f"Stored locally: {BIGNITION_CREDENTIALS_PATH}",
        ).grid(row=1, column=0, columnspan=7, sticky="w", pady=(8, 0))

        tools = ttk.Frame(container)
        tools.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(tools, text="Open Working Files", command=self.open_working_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(tools, text="Open Output Workbook", command=self.open_output_workbook).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(tools, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT)

        ttk.Label(container, textvariable=self.status_var).pack(anchor=tk.W, pady=(0, 6))

        log_frame = ttk.Frame(container)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def log(self, message: str) -> None:
        self.log_queue.put(f"[{timestamp()}] {message}")

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        self.root.after(120, self._drain_log_queue)

    def _set_running(self, running: bool, status_text: str) -> None:
        self.status_var.set(status_text)
        self.run_btn.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.setup_btn.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.pipeline_combo.configure(state=tk.DISABLED if running else "readonly")
        self.stop_btn.configure(state=tk.NORMAL if running else tk.DISABLED)

    def toggle_epic_password_visibility(self) -> None:
        self.epic_password_entry.configure(show="" if self.show_epic_password_var.get() else "*")

    def toggle_bignition_password_visibility(self) -> None:
        self.bignition_password_entry.configure(show="" if self.show_bignition_password_var.get() else "*")

    def load_saved_credentials(self) -> None:
        if EPIC_CREDENTIALS_PATH.exists():
            try:
                raw = json.loads(EPIC_CREDENTIALS_PATH.read_text(encoding="utf-8"))
                self.epic_usercode_var.set(str(raw.get("usercode", "")))
                self.epic_password_var.set(str(raw.get("password", "")))
                self.log("Loaded saved EPIC credentials.")
            except Exception as exc:
                self.log(f"Could not read saved EPIC credentials: {exc}")

        if BIGNITION_CREDENTIALS_PATH.exists():
            try:
                raw = json.loads(BIGNITION_CREDENTIALS_PATH.read_text(encoding="utf-8"))
                self.bignition_username_var.set(str(raw.get("username", "")))
                self.bignition_password_var.set(str(raw.get("password", "")))
                self.log("Loaded saved Bignition credentials.")
            except Exception as exc:
                self.log(f"Could not read saved Bignition credentials: {exc}")

    def save_credentials(self) -> None:
        usercode = self.epic_usercode_var.get().strip()
        password = self.epic_password_var.get().strip()
        if not usercode or not password:
            messagebox.showwarning("EPIC Credentials", "Usercode and Password are both required to save.")
            return
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "usercode": usercode,
            "password": password,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        EPIC_CREDENTIALS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        self.log("Saved EPIC credentials.")
        messagebox.showinfo("EPIC Credentials", "Credentials saved for EPIC auto-login.")

    def clear_credentials(self) -> None:
        self.epic_usercode_var.set("")
        self.epic_password_var.set("")
        try:
            if EPIC_CREDENTIALS_PATH.exists():
                EPIC_CREDENTIALS_PATH.unlink()
                self.log("Removed saved EPIC credentials.")
        except Exception as exc:
            self.log(f"Could not remove saved EPIC credentials: {exc}")
        messagebox.showinfo("EPIC Credentials", "Saved credentials cleared.")

    def save_bignition_credentials(self) -> None:
        username = self.bignition_username_var.get().strip()
        password = self.bignition_password_var.get().strip()
        if not username or not password:
            messagebox.showwarning("Bignition Credentials", "Username and Password are both required to save.")
            return
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "username": username,
            "password": password,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        BIGNITION_CREDENTIALS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        self.log("Saved Bignition credentials.")
        messagebox.showinfo("Bignition Credentials", "Credentials saved for Bignition auto-login.")

    def clear_bignition_credentials(self) -> None:
        self.bignition_username_var.set("")
        self.bignition_password_var.set("")
        try:
            if BIGNITION_CREDENTIALS_PATH.exists():
                BIGNITION_CREDENTIALS_PATH.unlink()
                self.log("Removed saved Bignition credentials.")
        except Exception as exc:
            self.log(f"Could not remove saved Bignition credentials: {exc}")
        messagebox.showinfo("Bignition Credentials", "Saved credentials cleared.")

    def start_setup_only(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self.stop_requested = False
        self._set_running(True, "Running setup...")
        self.worker_thread = threading.Thread(target=self._setup_worker, daemon=True)
        self.worker_thread.start()

    def start_pipeline(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self.stop_requested = False
        self._set_running(True, "Running pipeline...")
        self.worker_thread = threading.Thread(target=self._pipeline_worker, daemon=True)
        self.worker_thread.start()

    def stop_pipeline(self) -> None:
        self.stop_requested = True
        proc = self.current_process
        if proc and proc.poll() is None:
            self.log("Stop requested. Terminating current process...")
            try:
                subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], check=False, capture_output=True)
            except Exception as exc:
                self.log(f"Could not terminate process cleanly: {exc}")

    def _pipeline_worker(self) -> None:
        try:
            if self.setup_before_run_var.get():
                self.ensure_environment()
            if self.stop_requested:
                self.log("Pipeline cancelled before execution.")
                return

            pipeline_name = self.pipeline_var.get()
            steps = PIPELINES.get(pipeline_name, [])
            if not steps:
                raise RuntimeError(f"No steps configured for pipeline: {pipeline_name}")

            self.log(f"Starting pipeline: {pipeline_name}")
            for step in steps:
                if self.stop_requested:
                    self.log("Pipeline stopped by user.")
                    return
                self.run_python_script(step)
            self.log("Pipeline completed successfully.")
            self.log("Workflow is Complete")
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Pipeline Error", str(exc))
        finally:
            self.current_process = None
            self._set_running(False, "Idle")

    def _setup_worker(self) -> None:
        try:
            self.ensure_environment()
            self.log("Environment setup complete.")
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Setup Error", str(exc))
        finally:
            self.current_process = None
            self._set_running(False, "Idle")

    def ensure_environment(self) -> None:
        if not VENV_PYTHON.exists():
            bootstrap = self.find_bootstrap_python()
            if bootstrap is None:
                raise RuntimeError("No Python launcher found. Install Python 3 and retry.")
            self.log("Creating virtual environment...")
            self.run_command(bootstrap + ["-m", "venv", str(PROJECT_ROOT / ".venv")], label="Create venv")

        self.log("Upgrading pip...")
        self.run_command([str(VENV_PYTHON), "-m", "pip", "install", "--upgrade", "pip"], label="pip upgrade")

        if not REQUIREMENTS.exists():
            raise RuntimeError(f"Missing requirements file: {REQUIREMENTS}")
        self.log("Installing requirements...")
        self.run_command([str(VENV_PYTHON), "-m", "pip", "install", "-r", str(REQUIREMENTS)], label="pip install")

        pw_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright"
        has_chromium = pw_dir.exists() and any(pw_dir.glob("chromium-*"))
        if not has_chromium:
            self.log("Installing Playwright Chromium...")
            self.run_command([str(VENV_PYTHON), "-m", "playwright", "install", "chromium"], label="playwright install")
        else:
            self.log("Playwright Chromium already installed.")

    def find_bootstrap_python(self) -> list[str] | None:
        for candidate in (["py", "-3"], ["python"]):
            try:
                result = subprocess.run(
                    candidate + ["-c", "import sys"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=CREATE_NO_WINDOW,
                    timeout=10,
                    check=False,
                )
            except (OSError, subprocess.SubprocessError):
                continue
            if result.returncode == 0:
                return candidate

        for path in (
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python" / "Python312" / "python.exe",
            Path(os.environ.get("ProgramFiles", "")) / "Python312" / "python.exe",
        ):
            if path.exists():
                return [str(path)]
        return None

    def run_python_script(self, script_name: str) -> None:
        script_path = SCRIPTS_DIR / script_name
        if not script_path.exists():
            raise RuntimeError(f"Script not found: {script_path}")
        self.log(f"Running {script_name}...")
        self.run_command([str(VENV_PYTHON), str(script_path)], label=script_name)

    def run_command(self, args: list[str], label: str) -> None:
        if self.stop_requested:
            raise RuntimeError("Run stopped.")
        self.log(f"{label}: {' '.join(args)}")
        proc = subprocess.Popen(
            args,
            cwd=str(PROJECT_ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            creationflags=CREATE_NO_WINDOW,
        )
        self.current_process = proc
        assert proc.stdout is not None
        for line in proc.stdout:
            if line:
                self.log(line.rstrip())
        exit_code = proc.wait()
        self.current_process = None
        if self.stop_requested:
            raise RuntimeError("Run stopped.")
        if exit_code != 0:
            raise RuntimeError(f"{label} failed with exit code {exit_code}")

    def open_working_files(self) -> None:
        WORKING_FILES_DIR.mkdir(parents=True, exist_ok=True)
        os.startfile(str(WORKING_FILES_DIR))  # type: ignore[attr-defined]

    def open_output_workbook(self) -> None:
        if OUTPUT_WORKBOOK.exists():
            os.startfile(str(OUTPUT_WORKBOOK))  # type: ignore[attr-defined]
            return
        alt = OUTPUT_WORKBOOK.with_name(f"{OUTPUT_WORKBOOK.stem}.new{OUTPUT_WORKBOOK.suffix}")
        if alt.exists():
            os.startfile(str(alt))  # type: ignore[attr-defined]
            return
        messagebox.showinfo("Output Workbook", "Output workbook was not found yet.")

    def _on_close(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Exit", "A run is still active. Stop and exit?"):
                return
            self.stop_pipeline()
        self.root.destroy()


def main() -> int:
    root = tk.Tk()
    app = AutomationUI(root)
    app.log("UI ready.")
    app.log("Select a pipeline and click Run Selected.")
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
