from pathlib import Path
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import winreg

from openpyxl import load_workbook
from biophi_app import (
    IMAGE_SUFFIXES, Inputs, clean_name, measure_leaf, run_analysis, trial_info,
    write_error_log,
)
PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "template.pptx"
try:
    from brochure_maker import build_brochure
except Exception:
    build_brochure = None


def open_created_files(*paths):
    for path in paths:
        try:
            if path and Path(path).is_file():
                os.startfile(str(Path(path).resolve()))
        except OSError:
            pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # Read the persistent user value directly. Windows Explorer does not
        # always refresh inherited environment variables immediately.
        if not os.environ.get("ROBOFLOW_API_KEY", "").strip():
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as environment_key:
                    saved_key = winreg.QueryValueEx(environment_key, "ROBOFLOW_API_KEY")[0]
                if str(saved_key).strip():
                    os.environ["ROBOFLOW_API_KEY"] = str(saved_key).strip()
            except OSError:
                pass
        self.title("Biophi Analysis")
        self.geometry("820x520")
        self.minsize(760, 490)
        self.variables = {name: tk.StringVar() for name in (
            "master_folder", "trial_name", "evaluation", "environment",
            "shelf_life", "photos", "output_folder"
        )}
        self.status = tk.StringVar(value="Choose the master folder, review the predictions, then analyze.")
        self.roboflow_status = tk.StringVar()
        self.airtable_status = tk.StringVar()
        self._build()
        self._refresh_roboflow_status()
        self._refresh_airtable_status()
        self.bind("<FocusIn>", lambda _event: (self._refresh_roboflow_status(), self._refresh_airtable_status()))
        if not os.environ.get("ROBOFLOW_API_KEY", "").strip():
            self.after(400, self._set_roboflow_key)

    def _build(self):
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Biophi Analysis", font=("Segoe UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            frame,
            text="Unknown or unmeasurable values are left blank.",
            foreground="#4b5563",
        ).grid(row=1, column=0, sticky="w", pady=(0, 18))
        self.roboflow_label = ttk.Label(frame, textvariable=self.roboflow_status)
        self.roboflow_label.grid(row=1, column=1, columnspan=2, sticky="e", pady=(0, 18))

        ttk.Label(frame, text="Master folder").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(frame, textvariable=self.variables["master_folder"]).grid(
            row=2, column=1, sticky="ew", pady=6
        )
        ttk.Button(frame, text="Choose & predict…", command=self._choose_master).grid(
            row=2, column=2, padx=(10, 0), pady=6
        )
        fields = [
            ("Trial name", "trial_name", None),
            ("Evaluation workbook", "evaluation", False),
            ("Environment workbook", "environment", False),
            ("Shelf-life workbook (optional)", "shelf_life", False),
            ("Photos folder", "photos", True),
            ("Save results in", "output_folder", True),
        ]
        for row, (label, key, folder) in enumerate(fields, start=3):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=6)
            ttk.Entry(frame, textvariable=self.variables[key]).grid(
                row=row, column=1, sticky="ew", pady=6
            )
            if folder is not None:
                ttk.Button(
                    frame,
                    text="Change…",
                    command=lambda k=key, f=folder: self._browse(k, f),
                ).grid(row=row, column=2, padx=(10, 0), pady=6)

        frame.columnconfigure(1, weight=1)
        ttk.Separator(frame).grid(row=9, column=0, columnspan=3, sticky="ew", pady=(18, 12))
        ttk.Label(frame, textvariable=self.status).grid(row=10, column=0, sticky="w")
        roboflow_buttons = ttk.Frame(frame)
        roboflow_buttons.grid(row=10, column=1, sticky="e", padx=(10, 0))
        ttk.Button(roboflow_buttons, text="Set API Key", command=self._set_roboflow_key).pack(side="left")
        ttk.Button(roboflow_buttons, text="Test Roboflow", command=self._test_roboflow).pack(side="left", padx=(6, 0))
        self.run_button = ttk.Button(frame, text="Analyze", command=self._run)
        self.run_button.grid(row=10, column=2, sticky="e")
        airtable_frame = ttk.Frame(frame)
        airtable_frame.grid(row=11, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(airtable_frame, textvariable=self.airtable_status).pack(side="left")
        ttk.Button(airtable_frame, text="Set Airtable Token", command=self._set_airtable_token).pack(side="left", padx=(8, 0))
        ttk.Button(frame, text="Variety Overview from existing analysis", command=self._overview_existing).grid(
            row=11, column=2, sticky="e", pady=(8, 0)
        )

    def _overview_existing(self):
        analysis = filedialog.askopenfilename(
            title="Choose completed Biophi Analysis workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not analysis:
            return
        output_folder = filedialog.askdirectory(title="Choose overview output folder")
        if not output_folder:
            return
        if not build_brochure:
            messagebox.showerror("Overview unavailable", "The combined Variety Overview module could not be loaded.")
            return
        try:
            data_file, deck = build_brochure(Path(analysis), TEMPLATE_PATH, Path(output_folder), "")
            messagebox.showinfo("Overview complete", f"Saved Airtable data:\n{data_file}\n\nSaved overview:\n{deck}")
            open_created_files(data_file, deck)
        except Exception as error:
            messagebox.showerror("Overview failed", str(error))

    def _refresh_roboflow_status(self):
        key_detected = bool(os.environ.get("ROBOFLOW_API_KEY", "").strip())
        if key_detected:
            self.roboflow_status.set("Roboflow ready — Workflow Leafy will be used")
            self.roboflow_label.configure(foreground="#137333")
        else:
            self.roboflow_status.set("Roboflow cloud API key not found")
            self.roboflow_label.configure(foreground="#6b7280")

    def _refresh_airtable_status(self):
        token = os.environ.get("AIRTABLE_PAT", "").strip()
        if not token:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
                    token = str(winreg.QueryValueEx(key, "AIRTABLE_PAT")[0]).strip()
                    if token:
                        os.environ["AIRTABLE_PAT"] = token
            except OSError:
                pass
        cache = PROJECT_ROOT / "Assets" / "airtable_cache.json"
        if token:
            self.airtable_status.set("Airtable token saved — connected cache is used if live access is rejected")
        elif cache.is_file():
            self.airtable_status.set("Airtable connected cache ready")
        else:
            self.airtable_status.set("Airtable token not set — overview unavailable")

    def _set_airtable_token(self):
        token = simpledialog.askstring(
            "Set Airtable token",
            "Paste an Airtable Personal Access Token with read access to the Brochure Maker base. It is saved only in your Windows user environment:",
            show="*", parent=self,
        )
        if token is None:
            return
        token = token.strip()
        if not token:
            messagebox.showwarning("Empty token", "No Airtable token was saved.")
            return
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "AIRTABLE_PAT", 0, winreg.REG_SZ, token)
            os.environ["AIRTABLE_PAT"] = token
            self._refresh_airtable_status()
            messagebox.showinfo("Airtable token saved", "The token is active in this app. You can now create a brochure.")
        except OSError as error:
            messagebox.showerror("Could not save Airtable token", str(error))

    def _set_roboflow_key(self):
        key = simpledialog.askstring(
            "Set Roboflow API key",
            "Paste the private Roboflow API key. It will be stored in your Windows user environment:",
            show="*",
            parent=self,
        )
        if key is None:
            return
        key = key.strip()
        if not key:
            messagebox.showwarning("Empty API key", "No key was saved.")
            return
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                "Environment",
                0,
                winreg.KEY_SET_VALUE,
            ) as environment_key:
                winreg.SetValueEx(environment_key, "ROBOFLOW_API_KEY", 0, winreg.REG_SZ, key)
            os.environ["ROBOFLOW_API_KEY"] = key
            self._refresh_roboflow_status()
            messagebox.showinfo(
                "Roboflow key saved",
                "The key was saved for your Windows account and is active in this app. You can now click Test Roboflow.",
            )
        except OSError as error:
            messagebox.showerror("Could not save API key", str(error))

    def _test_roboflow(self):
        self._refresh_roboflow_status()
        if not os.environ.get("ROBOFLOW_API_KEY", "").strip():
            messagebox.showwarning(
                "Roboflow key not found",
                "Save ROBOFLOW_API_KEY as a Windows user environment variable, then close and reopen the app.",
            )
            return
        photo_folder = Path(self.variables["photos"].get())
        image = next(
            (path for path in photo_folder.rglob("*") if path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES),
            None,
        ) if photo_folder.is_dir() else None
        if image is None:
            messagebox.showwarning("No test photo", "Choose the master folder first so the app can test one photo.")
            return
        self.status.set("Testing Roboflow cloud connection…")
        self.update_idletasks()
        result = measure_leaf(image)
        if result and result.get("method") == "Roboflow Workflow Leafy":
            messagebox.showinfo(
                "Roboflow test successful",
                f"Workflow Leafy responded successfully.\nDetected leaf instances: {result['leaf_count']}",
            )
            self.status.set("Roboflow cloud test successful.")
        else:
            detail = ""
            if result and result.get("cloud_issues"):
                detail = "\n\n" + "\n".join(result["cloud_issues"])
            messagebox.showerror(
                "Roboflow test failed",
                "Workflow Leafy did not return a usable response. Check the API key, workflow access, and internet connection." + detail,
            )
            self.status.set("Roboflow cloud test failed.")

    def _choose_master(self):
        selected = filedialog.askdirectory(title="Choose trial master folder")
        if not selected:
            return
        master = Path(selected)
        self.variables["master_folder"].set(str(master))
        workbooks = list(master.rglob("*.xlsx")) + list(master.rglob("*.xlsm"))

        def predict_file(words):
            matches = [
                path for path in workbooks
                if all(word in path.name.lower() for word in words)
                and "output" not in path.name.lower()
            ]
            return matches[0] if matches else None

        evaluation = predict_file(["evaluation"])
        environment = predict_file(["environment"])
        shelf = predict_file(["shelf", "life"])
        photo_candidates = [
            path for path in master.rglob("*")
            if path.is_dir() and "photo" in path.name.lower()
        ]
        predictions = {
            "evaluation": evaluation,
            "environment": environment,
            "shelf_life": shelf,
            "photos": photo_candidates[0] if photo_candidates else None,
        }
        for key, path in predictions.items():
            if path:
                self.variables[key].set(str(path))
        output = master / "Biophi Analysis Output"
        self.variables["output_folder"].set(str(output))
        if evaluation:
            try:
                workbook = load_workbook(evaluation, read_only=False, data_only=True)
                name = clean_name(trial_info(workbook).get("trialname"))
                workbook.close()
                if name:
                    self.variables["trial_name"].set(name)
            except Exception:
                pass
        found = sum(path is not None for path in predictions.values())
        self.status.set(f"Predicted {found} of 4 inputs. Review or change any field before analyzing.")

    def _browse(self, key, folder):
        if folder:
            selected = filedialog.askdirectory(title="Choose folder")
        else:
            selected = filedialog.askopenfilename(
                title="Choose Excel workbook",
                filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
            )
        if selected:
            self.variables[key].set(selected)

    def _run(self):
        required = ("trial_name", "evaluation", "environment", "photos", "output_folder")
        missing = [key for key in required if not self.variables[key].get().strip()]
        if missing:
            messagebox.showwarning("Missing selections", "Please select every input and output location.")
            return
        self.run_button.configure(state="disabled")
        self.status.set("Analyzing data and photos…")
        self.update_idletasks()
        inputs = Inputs(
            evaluation=Path(self.variables["evaluation"].get()),
            environment=Path(self.variables["environment"].get()),
            shelf_life=Path(self.variables["shelf_life"].get()) if self.variables["shelf_life"].get().strip() else None,
            photos=Path(self.variables["photos"].get()),
            output_folder=Path(self.variables["output_folder"].get()),
            trial_name=self.variables["trial_name"].get(),
        )
        try:
            output, counts = run_analysis(inputs)
            self.status.set(f"Complete — {counts['varieties']} varieties written.")
            messagebox.showinfo(
                "Analysis complete",
                f"Saved:\n{output}\n\n"
                f"Harvest dates: {counts['harvests']}\n"
                f"Evaluation records: {counts['records']}\n"
                f"Leaf sizes measured: {counts['leaf_sizes']}\n"
                f"Photos requiring review: {counts['photo_reviews']}\n"
                f"Shelf-life values found: {counts['shelf_lives']}",
            )
            overview_created = False
            if build_brochure and messagebox.askyesno(
                "Create Variety Overview?",
                "Would you like to create the branded Variety Overview from this analysis and the Brochure Maker Airtable data?",
            ):
                try:
                    data_file, deck = build_brochure(output, TEMPLATE_PATH, inputs.output_folder, inputs.trial_name)
                    messagebox.showinfo("Overview complete", f"Saved Airtable data:\n{data_file}\n\nSaved overview:\n{deck}")
                    open_created_files(data_file, deck)
                    overview_created = True
                except Exception as overview_error:
                    messagebox.showerror("Overview failed", str(overview_error))
            if not overview_created:
                open_created_files(output)
        except Exception as error:
            try:
                log = write_error_log(inputs.output_folder, error)
                detail = f"\n\nDetails were saved to:\n{log}"
            except Exception:
                detail = ""
            self.status.set("Analysis failed.")
            messagebox.showerror("Analysis failed", f"{error}{detail}")
        finally:
            self.run_button.configure(state="normal")


if __name__ == "__main__":
    App().mainloop()
