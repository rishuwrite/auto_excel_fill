import os
import sys
import json
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "FBT Manifest Generator"
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fbt_settings.json")


def load_settings():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(output_folder):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"output_folder": output_folder}, f, indent=2)
    except Exception:
        pass


def resource_path(name):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def find_python():
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base, "runtime", "python.exe"),
        os.path.join(base, "python.exe"),
        sys.executable,
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return sys.executable


def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("760x500")
    root.minsize(700, 460)
    root.resizable(True, True)

    settings = load_settings()

    booking_var = tk.StringVar()
    mode_var = tk.StringVar(value="yes")
    output_var = tk.StringVar(
        value=settings.get(
            "output_folder",
            os.path.join(os.path.expanduser("~"), "Documents")
        )
    )
    status_var = tk.StringVar(value="Ready.")

    def update_button_state(*_):
        if booking_var.get().strip() and output_var.get().strip():
            generate_btn.config(state="normal")
        else:
            generate_btn.config(state="disabled")

    def choose_booking():
        path = filedialog.askopenfilename(
            title="Select Booking XLS",
            filetypes=[
                ("Excel files", "*.xls *.xlsx"),
                ("XLS files", "*.xls"),
                ("XLSX files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            booking_var.set(path)
            status_var.set("Booking file selected. Ready.")
            update_button_state()

    def choose_output():
        path = filedialog.askdirectory(title="Select FBT Output Folder")
        if path:
            output_var.set(path)
            save_settings(path)
            status_var.set("Output folder saved. It will be remembered next time.")
            update_button_state()

    def generate():
        booking = booking_var.get().strip()
        output = output_var.get().strip()
        mode = mode_var.get().strip().lower()

        if not booking:
            messagebox.showwarning("Missing booking file", "Please select the booking XLS file.")
            return

        if not os.path.isfile(booking):
            messagebox.showerror("Invalid booking file", "The selected booking file does not exist.")
            return

        if not output:
            messagebox.showwarning("Missing output folder", "Please select an output folder.")
            return

        try:
            os.makedirs(output, exist_ok=True)
            # The user selects/remembers only the BASE output folder.
            # Each run writes into a YYYY-MM-DD subfolder inside it.
            date_folder = os.path.join(
                output,
                __import__("datetime").date.today().isoformat()
            )
            os.makedirs(date_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Output folder error", str(e))
            return

        generator = resource_path("generate_fbt.py")
        if not os.path.isfile(generator):
            messagebox.showerror(
                "Missing file",
                "generate_fbt.py was not found beside this GUI."
            )
            return

        save_settings(output)
        generate_btn.config(state="disabled")
        status_var.set("Generating FBT files... Please wait.")
        root.update_idletasks()

        # generate_fbt.py requires both real template paths.
        base = os.path.dirname(os.path.abspath(__file__))
        suffix = "YES" if mode == "yes" else "NO"

        us_template = os.path.join(base, f"FBT_US_NFEI_{suffix}.xls")
        nonus_template = os.path.join(base, f"FBT_NON_US_NFEI_{suffix}.xls")

        # Keep compatibility with the existing bundled YES templates when
        # dedicated NFEI-NO templates have not been bundled yet.
        if not os.path.isfile(us_template):
            us_template = os.path.join(base, "FBT_US_NFEI_YES.xls")
        if not os.path.isfile(nonus_template):
            nonus_template = os.path.join(base, "FBT_NON_US_NFEI_YES.xls")

        if not os.path.isfile(us_template) or not os.path.isfile(nonus_template):
            messagebox.showerror(
                "Templates missing",
                "The required FBT template files were not found beside the application."
            )
            generate_btn.config(state="normal")
            return

        python_exe = find_python()
        cmd = [
            python_exe,
            generator,
            booking,
            "--mode", mode,
            "--us-template", us_template,
            "--nonus-template", nonus_template,
            "--outdir", date_folder,
        ]

        try:
            result = subprocess.run(
                cmd,
                cwd=os.path.dirname(os.path.abspath(__file__)),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            if result.returncode == 0:
                status_var.set(
                    f"Generation completed successfully. Files saved in: {date_folder}"
                )
                details = (result.stdout or "").strip()
                messagebox.showinfo(
                    "FBT Generation Complete",
                    f"FBT files were created successfully.\n\n"
                    f"Output folder:\n{date_folder}\n\n" +
                    (details[-3000:] if details else "")
                )
            else:
                status_var.set("Generation failed.")
                details = ((result.stderr or "") + "\n" + (result.stdout or "")).strip()
                messagebox.showerror(
                    "Generation Failed",
                    details[-6000:] if details else "The generator returned an error."
                )

        except Exception as e:
            status_var.set("Generation failed.")
            messagebox.showerror("Generation Failed", str(e))
        finally:
            update_button_state()

    # ---------- GUI ----------
    main_frame = ttk.Frame(root, padding=24)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(
        main_frame, text=APP_TITLE, font=("Segoe UI", 18, "bold")
    ).pack(anchor="w", pady=(0, 8))

    ttk.Label(
        main_frame,
        text="Select the booking export. The correct US / Non-US files will be created automatically.",
        wraplength=690,
    ).pack(anchor="w", pady=(0, 22))

    booking_frame = ttk.Frame(main_frame)
    booking_frame.pack(fill="x", pady=7)
    ttk.Label(booking_frame, text="Booking .xls:", width=16).pack(side="left")
    ttk.Entry(booking_frame, textvariable=booking_var).pack(
        side="left", fill="x", expand=True, padx=(0, 10)
    )
    ttk.Button(
        booking_frame, text="Browse...", command=choose_booking, width=12
    ).pack(side="right")

    mode_frame = ttk.Frame(main_frame)
    mode_frame.pack(fill="x", pady=7)
    ttk.Label(mode_frame, text="NFEI mode:", width=16).pack(side="left")
    ttk.Combobox(
        mode_frame,
        textvariable=mode_var,
        values=["yes", "no"],
        state="readonly",
        width=10,
    ).pack(side="left")
    ttk.Label(mode_frame, text="  YES / NO mode").pack(side="left", padx=8)

    output_frame = ttk.Frame(main_frame)
    output_frame.pack(fill="x", pady=7)
    ttk.Label(output_frame, text="Output folder:", width=16).pack(side="left")
    ttk.Entry(output_frame, textvariable=output_var).pack(
        side="left", fill="x", expand=True, padx=(0, 10)
    )
    ttk.Button(
        output_frame, text="Change...", command=choose_output, width=12
    ).pack(side="right")

    ttk.Label(
        main_frame,
        text="First use: choose a folder. Next use: the same folder is remembered automatically. "
             "A YYYY-MM-DD date folder is created inside it.",
        wraplength=690,
    ).pack(anchor="w", pady=(8, 18))

    status_box = ttk.LabelFrame(main_frame, text="Status", padding=12)
    status_box.pack(fill="x", pady=(0, 18))
    ttk.Label(status_box, textvariable=status_var, wraplength=650).pack(anchor="w")

    button_area = ttk.Frame(main_frame)
    button_area.pack(fill="x", pady=(4, 0))

    generate_btn = ttk.Button(
        button_area,
        text="GENERATE FBT FILES",
        command=generate,
        state="disabled",
    )
    generate_btn.pack(ipadx=28, ipady=10)

    ttk.Label(
        main_frame,
        text="No administrator rights are normally required when Microsoft Excel is already installed.",
        wraplength=690,
    ).pack(anchor="w", pady=(18, 0))

    booking_var.trace_add("write", update_button_state)
    output_var.trace_add("write", update_button_state)
    update_button_state()

    root.mainloop()


if __name__ == "__main__":
    main()
