import os
from decimal import Decimal, ROUND_HALF_UP
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl


def normalize_id(val):
    """Normalize a Vendor ID (handles ints, floats such as 5318335.0, and strings with spaces)."""
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s if s else None


def to_float(val):
    """Safely convert a cell value to float (returns None if it is blank or not a number)."""
    if val is None or isinstance(val, bool):
        return None
    try:
        s = str(val).strip()
        if s == "":
            return None
        return float(s)
    except ValueError:
        return None


def round_half_up(x, digits):
    """Round the way Excel does (0.5 always rounds up).
    Python's built-in round() uses banker's rounding, e.g. round(4.25, 1) gives 4.2,
    while this function gives 4.3. Returns an int when digits == 0."""
    d = Decimal(str(round(x, 10)))  # round first to remove floating-point noise
    q = d.quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)
    return int(q) if digits == 0 else float(q)


class VendorUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vendor Score Updater (QF 68B -> QF 01C)")
        self.root.geometry("680x470")
        self.root.resizable(False, False)

        self.file_68b = tk.StringVar()
        self.file_01c = tk.StringVar()
        self.round_digits = tk.IntVar(value=1)

        self._build_ui()

    def _build_ui(self):
        pad_opts = {"padx": 12, "pady": 6}

        # 1. Select the QF 68B file
        frame_1 = tk.LabelFrame(self.root, text="Step 1: Select the source file (QF 68B.xlsx)", padx=10, pady=8)
        frame_1.pack(fill="x", **pad_opts)

        tk.Entry(frame_1, textvariable=self.file_68b, width=65, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(frame_1, text="Browse...", width=10, command=self.select_68b).pack(side="right")

        # 2. Select the QF 01C file
        frame_2 = tk.LabelFrame(self.root, text="Step 2: Select the target file (QF 01C-Approved Vendors List.xlsx)", padx=10, pady=8)
        frame_2.pack(fill="x", **pad_opts)

        tk.Entry(frame_2, textvariable=self.file_01c, width=65, state="readonly").pack(side="left", padx=(0, 8))
        tk.Button(frame_2, text="Browse...", width=10, command=self.select_01c).pack(side="right")

        # 3. Options
        frame_opt = tk.LabelFrame(self.root, text="Options", padx=10, pady=8)
        frame_opt.pack(fill="x", **pad_opts)

        sub_opt = tk.Frame(frame_opt)
        sub_opt.pack(anchor="w")
        tk.Label(sub_opt, text="Decimal places for averages:").pack(side="left")
        tk.Spinbox(sub_opt, from_=0, to=6, textvariable=self.round_digits, width=5).pack(side="left")

        # 4. Run button (same color as the window background)
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=12, pady=8)

        self.run_btn = tk.Button(
            btn_frame,
            text="Calculate and Generate Updated Excel File",
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            height=2,
            command=self.process_files
        )
        self.run_btn.pack(fill="x")

        # 5. Log area
        log_frame = tk.LabelFrame(self.root, text="Log", padx=8, pady=6)
        log_frame.pack(fill="both", expand=True, **pad_opts)

        self.log_box = scrolledtext.ScrolledText(log_frame, height=11, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True)

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

    def select_68b(self):
        path = filedialog.askopenfilename(
            title="Select the Copy of QF 68B file",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if path:
            self.file_68b.set(path)

    def select_01c(self):
        path = filedialog.askopenfilename(
            title="Select the QF 01C Approved Vendors List file",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if path:
            self.file_01c.set(path)

    def process_files(self):
        path_68b = self.file_68b.get()
        path_01c = self.file_01c.get()

        if not path_68b or not path_01c:
            messagebox.showwarning("Notice", "Please select both Excel files first.")
            return

        self.log_box.delete("1.0", tk.END)
        self.log("Reading QF 68B data...")

        try:
            # 1. Read QF 68B
            wb_68b = openpyxl.load_workbook(path_68b, data_only=True)
            ws_68b = wb_68b.active

            # Structure: { vendor_id: {"count": int, "sums": [N, O, P, Q, R]} }
            vendor_stats = {}

            # Start from row 2 (row 1 is the header)
            # Column L = 12 (Vendor ID); N = 14, O = 15, P = 16, Q = 17, R = 18 (scores)
            for row_idx in range(2, ws_68b.max_row + 1):
                raw_vid = ws_68b.cell(row=row_idx, column=12).value  # column L
                vid = normalize_id(raw_vid)
                if not vid or vid.lower() == "vendor id":
                    continue

                if vid not in vendor_stats:
                    vendor_stats[vid] = {"count": 0, "sums": [0.0, 0.0, 0.0, 0.0, 0.0]}

                vendor_stats[vid]["count"] += 1

                # Blank or non-numeric scores are counted as 0 (they add nothing to the sum
                # but the row still counts toward the divisor)
                for idx, col_num in enumerate([14, 15, 16, 17, 18]):  # N, O, P, Q, R
                    val = to_float(ws_68b.cell(row=row_idx, column=col_num).value)
                    if val is not None:
                        vendor_stats[vid]["sums"][idx] += val

            if not vendor_stats:
                self.log("WARNING: No valid Vendor ID found in column L of QF 68B!")
                messagebox.showwarning("No valid data", "No Vendor ID was found in column L of QF 68B.")
                return

            digits = self.round_digits.get()
            vendor_avgs = {}
            self.log(f"Found {len(vendor_stats)} distinct Vendor ID(s):")
            for vid, info in vendor_stats.items():
                cnt = info["count"]
                avgs = [round_half_up(s / cnt, digits) for s in info["sums"]]
                vendor_avgs[vid] = avgs
                self.log(f"  - Vendor ID [{vid}] (appears {cnt} time(s)) -> averages (N~R): {avgs}")

            # 2. Open and update QF 01C
            self.log("\nLoading the QF 01C target file...")
            wb_01c = openpyxl.load_workbook(path_01c)
            ws_01c = wb_01c["Database"] if "Database" in wb_01c.sheetnames else wb_01c.active

            matched_vids = set()
            updated_rows = 0

            # Column J = 10 (Vendor ID); M = 13, N = 14, O = 15, P = 16, Q = 17, R = 18
            for row_idx in range(12, ws_01c.max_row + 1):
                raw_j = ws_01c.cell(row=row_idx, column=10).value  # column J
                vid = normalize_id(raw_j)
                if not vid or vid.lower() == "vendor id":
                    continue

                if vid in vendor_avgs:
                    avgs = vendor_avgs[vid]
                    # Write M, N, O, P, Q (columns 13~17); existing values are overwritten
                    for idx, target_col in enumerate([13, 14, 15, 16, 17]):
                        ws_01c.cell(row=row_idx, column=target_col).value = avgs[idx]

                    # Write the total into column R (column 18) as a plain value
                    # (this replaces any existing formula in that cell)
                    total_val = round_half_up(sum(avgs), digits)
                    ws_01c.cell(row=row_idx, column=18).value = total_val

                    matched_vids.add(vid)
                    updated_rows += 1
                    self.log(f"  OK  Row {row_idx} updated (Vendor ID: {vid}) -> M~Q: {avgs}, R total: {total_val}")

            # Report Vendor IDs that were not found in the target file
            unmatched = set(vendor_avgs.keys()) - matched_vids
            if unmatched:
                self.log(f"\nWARNING: These Vendor IDs were not found in column J of QF 01C: {', '.join(sorted(unmatched))}")

            # 3. Save as a new file so the original is never overwritten
            base_dir = os.path.dirname(path_01c)
            base_name, ext = os.path.splitext(os.path.basename(path_01c))
            output_path = os.path.join(base_dir, f"{base_name}_Updated{ext}")

            wb_01c.save(output_path)
            self.log(f"\nDone! {updated_rows} row(s) updated.")
            self.log(f"Saved to: {output_path}")
            messagebox.showinfo("Done", f"Done! {updated_rows} row(s) updated.\n\nSaved as:\n{output_path}")

        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while processing:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VendorUpdaterApp(root)
    root.mainloop()