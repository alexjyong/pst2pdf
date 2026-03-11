#!/usr/bin/env python3
"""pst2pdf GUI — CustomTkinter wrapper around the pst2pdf converter."""

from __future__ import annotations

import logging
import queue
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from parser import parse_pst
from renderer import render_email_to_pdf
from pypdf import PdfWriter
import csv

# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

_LOG_QUEUE: queue.Queue = queue.Queue()


class QueueHandler(logging.Handler):
    """Logging handler that puts records onto a queue for the GUI to consume."""
    def emit(self, record):
        _LOG_QUEUE.put(self.format(record))


def _setup_logging():
    handler = QueueHandler()
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.addHandler(handler)
    root_logger.setLevel(logging.INFO)


# ── Main window ───────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("pst2pdf")
        self.geometry("620x780")
        self.resizable(True, True)
        self.minsize(560, 680)

        self._running = False
        self._thread: threading.Thread | None = None

        self._build_ui()
        self._poll_log()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        row = 0

        # ── Title
        ctk.CTkLabel(self, text="pst2pdf", font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=row, column=0, padx=20, pady=(18, 2), sticky="w"
        )
        row += 1
        ctk.CTkLabel(self, text="Convert PST mailboxes to PDF for e-discovery",
                     text_color="gray").grid(row=row, column=0, padx=20, pady=(0, 12), sticky="w")
        row += 1

        # ── Basic options frame
        basic = ctk.CTkFrame(self)
        basic.grid(row=row, column=0, padx=16, pady=4, sticky="ew")
        basic.grid_columnconfigure(1, weight=1)
        row += 1

        # PST file
        ctk.CTkLabel(basic, text="PST File").grid(row=0, column=0, padx=12, pady=8, sticky="w")
        self._pst_var = ctk.StringVar()
        ctk.CTkEntry(basic, textvariable=self._pst_var, placeholder_text="Select a .pst file…").grid(
            row=0, column=1, padx=6, pady=8, sticky="ew"
        )
        ctk.CTkButton(basic, text="Browse", width=72, command=self._browse_pst).grid(
            row=0, column=2, padx=(0, 12), pady=8
        )

        # Output directory
        ctk.CTkLabel(basic, text="Output Dir").grid(row=1, column=0, padx=12, pady=8, sticky="w")
        self._out_var = ctk.StringVar()
        ctk.CTkEntry(basic, textvariable=self._out_var, placeholder_text="Select output folder…").grid(
            row=1, column=1, padx=6, pady=8, sticky="ew"
        )
        ctk.CTkButton(basic, text="Browse", width=72, command=self._browse_output).grid(
            row=1, column=2, padx=(0, 12), pady=8
        )

        # Prefix + Attachments side by side
        ctk.CTkLabel(basic, text="Prefix").grid(row=2, column=0, padx=12, pady=8, sticky="w")
        inner = ctk.CTkFrame(basic, fg_color="transparent")
        inner.grid(row=2, column=1, columnspan=2, padx=6, pady=8, sticky="ew")
        inner.grid_columnconfigure(1, weight=1)

        self._prefix_var = ctk.StringVar(value="MSG")
        ctk.CTkEntry(inner, textvariable=self._prefix_var, width=100).grid(
            row=0, column=0, padx=(0, 16), sticky="w"
        )
        ctk.CTkLabel(inner, text="Attachments").grid(row=0, column=1, padx=(0, 6), sticky="e")
        self._att_var = ctk.StringVar(value="embed")
        ctk.CTkOptionMenu(inner, variable=self._att_var, values=["embed", "extract", "none"],
                          width=110).grid(row=0, column=2, padx=(0, 12), sticky="e")

        # Manifest + Merge
        checks = ctk.CTkFrame(basic, fg_color="transparent")
        checks.grid(row=3, column=0, columnspan=3, padx=12, pady=(4, 10), sticky="ew")

        self._manifest_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(checks, text="Write manifest.csv", variable=self._manifest_var).pack(
            side="left", padx=(0, 20)
        )
        self._merge_var = ctk.BooleanVar(value=False)
        merge_cb = ctk.CTkCheckBox(checks, text="Merge to:", variable=self._merge_var,
                                   command=self._toggle_merge)
        merge_cb.pack(side="left", padx=(0, 6))
        self._merge_name_var = ctk.StringVar(value="all.pdf")
        self._merge_entry = ctk.CTkEntry(checks, textvariable=self._merge_name_var, width=100,
                                         state="disabled")
        self._merge_entry.pack(side="left")

        row += 1

        # ── Advanced collapsible
        self._adv_open = ctk.BooleanVar(value=False)
        self._adv_btn = ctk.CTkButton(
            self, text="▶  Advanced options", fg_color="transparent",
            text_color=("gray30", "gray70"), hover=False, anchor="w",
            command=self._toggle_advanced,
        )
        self._adv_btn.grid(row=row, column=0, padx=16, pady=(8, 0), sticky="ew")
        row += 1

        self._adv_frame = ctk.CTkFrame(self)
        self._adv_frame.grid_columnconfigure(1, weight=1)
        self._adv_row = row
        row += 1

        self._build_advanced(self._adv_frame)
        # Hidden by default — don't grid it yet

        # ── Convert button
        self._convert_btn = ctk.CTkButton(
            self, text="Convert", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_convert,
        )
        self._convert_btn.grid(row=row, column=0, padx=16, pady=12, sticky="ew")
        row += 1

        # ── Progress bar
        self._progress = ctk.CTkProgressBar(self)
        self._progress.set(0)
        self._progress.grid(row=row, column=0, padx=16, pady=(0, 8), sticky="ew")
        row += 1

        # ── Log output
        ctk.CTkLabel(self, text="Log", anchor="w").grid(
            row=row, column=0, padx=16, pady=(4, 0), sticky="w"
        )
        row += 1
        self._log_box = ctk.CTkTextbox(self, height=200, font=ctk.CTkFont(family="Courier", size=11))
        self._log_box.grid(row=row, column=0, padx=16, pady=(2, 16), sticky="nsew")
        self.grid_rowconfigure(row, weight=1)

    def _build_advanced(self, parent: ctk.CTkFrame):
        r = 0

        def row_pair(label, widget_fn, **kw):
            nonlocal r
            ctk.CTkLabel(parent, text=label).grid(row=r, column=0, padx=12, pady=6, sticky="w")
            widget_fn(parent, **kw).grid(row=r, column=1, padx=(6, 12), pady=6, sticky="w")
            r += 1

        # Start number
        ctk.CTkLabel(parent, text="Start Number").grid(row=r, column=0, padx=12, pady=6, sticky="w")
        self._start_var = ctk.StringVar(value="1")
        ctk.CTkEntry(parent, textvariable=self._start_var, width=80).grid(
            row=r, column=1, padx=(6, 12), pady=6, sticky="w"
        )
        r += 1

        # Max attachment size
        ctk.CTkLabel(parent, text="Max Attachment (MB)").grid(row=r, column=0, padx=12, pady=6, sticky="w")
        self._max_att_var = ctk.StringVar(value="10")
        ctk.CTkEntry(parent, textvariable=self._max_att_var, width=80).grid(
            row=r, column=1, padx=(6, 12), pady=6, sticky="w"
        )
        r += 1

        # Bates stamp
        self._bates_var = ctk.BooleanVar(value=False)
        bates_row = ctk.CTkFrame(parent, fg_color="transparent")
        bates_row.grid(row=r, column=0, columnspan=2, padx=12, pady=6, sticky="w")
        ctk.CTkCheckBox(bates_row, text="Bates stamp", variable=self._bates_var).pack(side="left", padx=(0, 16))
        ctk.CTkLabel(bates_row, text="Pad width:").pack(side="left", padx=(0, 6))
        self._bates_pad_var = ctk.StringVar(value="6")
        ctk.CTkEntry(bates_row, textvariable=self._bates_pad_var, width=50).pack(side="left")
        r += 1

        # Dedup + non-email
        flags = ctk.CTkFrame(parent, fg_color="transparent")
        flags.grid(row=r, column=0, columnspan=2, padx=12, pady=(4, 10), sticky="w")
        self._nodedup_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(flags, text="Disable deduplication", variable=self._nodedup_var).pack(
            side="left", padx=(0, 20)
        )
        self._nonemail_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(flags, text="Include non-email items", variable=self._nonemail_var).pack(side="left")
        r += 1

    # ── Actions ───────────────────────────────────────────────────────────────

    def _browse_pst(self):
        path = filedialog.askopenfilename(
            title="Select PST file",
            filetypes=[("PST files", "*.pst"), ("All files", "*.*")],
        )
        if path:
            self._pst_var.set(path)
            if not self._out_var.get():
                self._out_var.set(str(Path(path).parent / "output"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output directory")
        if path:
            self._out_var.set(path)

    def _toggle_merge(self):
        self._merge_entry.configure(state="normal" if self._merge_var.get() else "disabled")

    def _toggle_advanced(self):
        if self._adv_open.get():
            self._adv_frame.grid_forget()
            self._adv_open.set(False)
            self._adv_btn.configure(text="▶  Advanced options")
        else:
            self._adv_frame.grid(row=self._adv_row, column=0, padx=16, pady=(0, 4), sticky="ew")
            self._adv_open.set(True)
            self._adv_btn.configure(text="▼  Advanced options")

    def _on_convert(self):
        if self._running:
            return

        pst = self._pst_var.get().strip()
        out = self._out_var.get().strip()

        if not pst or not Path(pst).is_file():
            messagebox.showerror("Error", "Please select a valid PST file.")
            return
        if not out:
            messagebox.showerror("Error", "Please select an output directory.")
            return

        try:
            start_num = int(self._start_var.get())
            max_att = float(self._max_att_var.get())
            bates_pad = int(self._bates_pad_var.get())
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid value: {e}")
            return

        opts = {
            "pst_path": pst,
            "output_dir": out,
            "prefix": self._prefix_var.get().strip() or "MSG",
            "start_num": start_num,
            "attachments": self._att_var.get(),
            "manifest": self._manifest_var.get(),
            "merge": self._merge_name_var.get().strip() if self._merge_var.get() else None,
            "bates": self._bates_var.get(),
            "bates_pad": bates_pad,
            "max_attachment_mb": max_att,
            "no_dedup": self._nodedup_var.get(),
            "include_non_email": self._nonemail_var.get(),
        }

        self._log_box.delete("1.0", "end")
        self._progress.set(0)
        self._progress.configure(mode="indeterminate")
        self._progress.start()
        self._convert_btn.configure(state="disabled", text="Converting…")
        self._running = True

        self._thread = threading.Thread(target=self._run_conversion, args=(opts,), daemon=True)
        self._thread.start()

    def _run_conversion(self, opts: dict):
        try:
            _run(opts)
        except Exception as exc:
            logging.getLogger().error("Unexpected error: %s", exc)
        finally:
            _LOG_QUEUE.put("__DONE__")

    # ── Log polling ───────────────────────────────────────────────────────────

    def _poll_log(self):
        try:
            while True:
                msg = _LOG_QUEUE.get_nowait()
                if msg == "__DONE__":
                    self._progress.stop()
                    self._progress.configure(mode="determinate")
                    self._progress.set(1)
                    self._convert_btn.configure(state="normal", text="Convert")
                    self._running = False
                else:
                    self._log_box.insert("end", msg + "\n")
                    self._log_box.see("end")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)


# ── Conversion logic (runs in worker thread) ──────────────────────────────────

def _run(opts: dict):
    logger = logging.getLogger(__name__)

    pst_path = opts["pst_path"]
    output_dir = Path(opts["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)

    prefix = opts["prefix"]
    start_num = opts["start_num"]
    attachments = opts["attachments"]
    bates = opts["bates"]
    bates_pad = opts["bates_pad"]
    max_att_bytes = int(opts["max_attachment_mb"] * 1024 * 1024) if opts["max_attachment_mb"] > 0 else 0
    parse_att_bytes = -1 if attachments == "none" else max_att_bytes

    manifest_rows: list[dict] = []
    output_files: list[Path] = []
    counter = start_num
    bates_counter = start_num
    seen_ids: set[str] = set()
    success = skipped = deduped = 0

    logger.info("Parsing %s …", pst_path)

    for email in parse_pst(pst_path, email_only=not opts["include_non_email"],
                            max_attachment_bytes=parse_att_bytes):
        seq = f"{counter:05d}"
        filename = f"{prefix}_{seq}.pdf"
        output_path = output_dir / filename

        if not opts["no_dedup"] and email.message_id:
            if email.message_id in seen_ids:
                deduped += 1
                continue
            seen_ids.add(email.message_id)

        bates_first = bates_counter
        try:
            bates_next = render_email_to_pdf(
                email, str(output_path),
                bates_prefix=prefix if bates else "",
                bates_start=bates_counter,
                bates_pad_width=bates_pad,
                embed_attachments=(attachments == "embed"),
            )
            if bates:
                bates_counter = bates_next
                logger.info("[%s] pages %s–%s  %s", filename,
                            f"{bates_first:0{bates_pad}d}", f"{bates_next - 1:0{bates_pad}d}",
                            email.subject or "(No Subject)")
            else:
                logger.info("[%s] %s", filename, email.subject or "(No Subject)")

            if attachments == "extract":
                for i, att in enumerate(email.attachments):
                    if att.is_embedded_msg or not att.data:
                        continue
                    safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in att.name).strip()
                    (output_dir / f"{Path(filename).stem}_{i + 1:02d}_{safe}").write_bytes(att.data)

            if opts["manifest"]:
                row = {
                    "filename": filename,
                    "date": email.date.isoformat() if email.date else "",
                    "from": email.from_, "to": "; ".join(email.to), "cc": "; ".join(email.cc),
                    "subject": email.subject, "message_id": email.message_id,
                    "folder_path": email.folder_path,
                    "attachments": "; ".join(a.name for a in email.attachments),
                }
                if bates:
                    row["bates_begin"] = f"{prefix}_{bates_first:0{bates_pad}d}"
                    row["bates_end"] = f"{prefix}_{bates_next - 1:0{bates_pad}d}"
                manifest_rows.append(row)

            output_files.append(output_path)
            success += 1
        except Exception as exc:
            logger.warning("Failed to render %s: %s", filename, exc)
            skipped += 1

        counter += 1

    if opts["manifest"] and manifest_rows:
        fieldnames = ["filename", "date", "from", "to", "cc", "subject",
                      "message_id", "folder_path", "attachments"]
        if bates:
            fieldnames += ["bates_begin", "bates_end"]
        with open(output_dir / "manifest.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(manifest_rows)
        logger.info("Manifest written.")

    if opts["merge"] and output_files:
        merge_path = output_dir / opts["merge"]
        logger.info("Merging %d PDFs → %s …", len(output_files), opts["merge"])
        writer = PdfWriter()
        for pdf in output_files:
            writer.append(str(pdf))
        with open(merge_path, "wb") as f:
            writer.write(f)
        logger.info("Merge complete.")

    if deduped:
        logger.info("Deduplicated: %d skipped.", deduped)
    logger.info("Done. %d/%d converted. %d skipped.", success, success + skipped, skipped)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    _setup_logging()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
