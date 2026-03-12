#!/usr/bin/env python3
"""pst2pdf GUI — CustomTkinter wrapper around the pst2pdf converter."""

from __future__ import annotations

import logging
import queue
import threading
from pathlib import Path

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox

from parser import parse_pst
from renderer import render_email_to_pdf
from pypdf import PdfWriter
from version import __version__
import csv
import sys, subprocess, os

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

_LOG_QUEUE: queue.Queue = queue.Queue()


class Tooltip:
    """Hover tooltip for any tkinter/CustomTkinter widget."""

    def __init__(self, widget, text: str, delay: int = 500):
        self._widget = widget
        self._text = text
        self._delay = delay
        self._tip: tk.Toplevel | None = None
        self._after_id: str | None = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._cancel)
        widget.bind("<ButtonPress>", self._cancel)

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self._widget.after(self._delay, self._show)

    def _cancel(self, _event=None):
        if self._after_id:
            self._widget.after_cancel(self._after_id)
            self._after_id = None
        self._hide()

    def _show(self):
        if self._tip:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self._text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("TkDefaultFont", 9), wraplength=300, padx=6, pady=4,
        ).pack()

    def _hide(self):
        if self._tip:
            self._tip.destroy()
            self._tip = None


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


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"pst2pdf {__version__}")
        self.geometry("620x700")

        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2) - 250
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.resizable(True, True)
        self.minsize(560, 680)

        self._running = False
        self._thread: threading.Thread | None = None

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        row = 0

        ctk.CTkLabel(self, text="pst2pdf", font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=row, column=0, padx=20, pady=(18, 2), sticky="w"
        )
        row += 1
        ctk.CTkLabel(self, text="Convert PST mailboxes to PDF for e-discovery",
                     text_color="gray").grid(row=row, column=0, padx=20, pady=(0, 12), sticky="w")
        row += 1
        row += 1
        ctk.CTkLabel(self, text="Mouse over any label to get a tooltip explaining it",
                     text_color="gray").grid(row=row, column=0, padx=20, pady=(0, 12), sticky="w")
        row += 1

        basic = ctk.CTkFrame(self)
        basic.grid(row=row, column=0, padx=16, pady=4, sticky="ew")
        basic.grid_columnconfigure(1, weight=1)
        row += 1

        pst_lbl = ctk.CTkLabel(basic, text="PST File")
        pst_lbl.grid(row=0, column=0, padx=12, pady=8, sticky="w")
        Tooltip(pst_lbl, "The Outlook .pst mailbox file to convert to PDF.")
        self._pst_var = ctk.StringVar()
        pst_entry = ctk.CTkEntry(basic, textvariable=self._pst_var, placeholder_text="Select a .pst file…")
        pst_entry.grid(row=0, column=1, padx=6, pady=8, sticky="ew")
        Tooltip(pst_entry, "The Outlook .pst mailbox file to convert to PDF.")
        ctk.CTkButton(basic, text="Browse", width=72, command=self._browse_pst).grid(
            row=0, column=2, padx=(0, 12), pady=8
        )

        out_lbl = ctk.CTkLabel(basic, text="Output Dir")
        out_lbl.grid(row=1, column=0, padx=12, pady=8, sticky="w")
        Tooltip(out_lbl, "Folder where output PDFs (and optional manifest) will be saved.\nCreated automatically if it doesn't exist.")
        self._out_var = ctk.StringVar()
        out_entry = ctk.CTkEntry(basic, textvariable=self._out_var, placeholder_text="Select output folder…")
        out_entry.grid(row=1, column=1, padx=6, pady=8, sticky="ew")
        Tooltip(out_entry, "Folder where output PDFs (and optional manifest) will be saved.\nCreated automatically if it doesn't exist.")
        ctk.CTkButton(basic, text="Browse", width=72, command=self._browse_output).grid(
            row=1, column=2, padx=(0, 12), pady=8
        )

        prefix_lbl = ctk.CTkLabel(basic, text="Filename Prefix")
        prefix_lbl.grid(row=2, column=0, padx=12, pady=8, sticky="w")
        _prefix_tip = "Text prepended to every output filename (and Bates number if enabled).\nExample: 'MSG' → MSG_00001.pdf, MSG_00002.pdf, …"
        Tooltip(prefix_lbl, _prefix_tip)
        inner = ctk.CTkFrame(basic, fg_color="transparent")
        inner.grid(row=2, column=1, columnspan=2, padx=6, pady=8, sticky="ew")
        inner.grid_columnconfigure(1, weight=1)

        self._prefix_var = ctk.StringVar(value="MSG")
        prefix_entry = ctk.CTkEntry(inner, textvariable=self._prefix_var, width=100,
                                    placeholder_text="e.g. MSG")
        prefix_entry.grid(row=0, column=0, padx=(0, 16), sticky="w")
        Tooltip(prefix_entry, _prefix_tip)
        att_lbl = ctk.CTkLabel(inner, text="Attachments")
        att_lbl.grid(row=0, column=1, padx=(0, 6), sticky="e")
        _att_tip = ("How to handle email attachments:\n"
                    "  embed   – append them inside each PDF\n"
                    "  extract – save as separate files alongside the PDF\n"
                    "  none    – ignore attachments entirely")
        Tooltip(att_lbl, _att_tip)
        self._att_var = ctk.StringVar(value="embed")
        att_menu = ctk.CTkOptionMenu(inner, variable=self._att_var, values=["embed", "extract", "none"],
                                     width=110)
        att_menu.grid(row=0, column=2, padx=(0, 12), sticky="e")
        Tooltip(att_menu, _att_tip)

        checks = ctk.CTkFrame(basic, fg_color="transparent")
        checks.grid(row=3, column=0, columnspan=3, padx=12, pady=(4, 10), sticky="ew")

        self._manifest_var = ctk.BooleanVar(value=True)
        manifest_cb = ctk.CTkCheckBox(checks, text="Write manifest.csv", variable=self._manifest_var)
        manifest_cb.pack(side="left", padx=(0, 20))
        Tooltip(manifest_cb, "Save a CSV index of all converted emails with metadata\n(date, sender, recipients, subject, folder path, attachments).")
        self._merge_var = ctk.BooleanVar(value=False)
        merge_cb = ctk.CTkCheckBox(checks, text="Merge to:", variable=self._merge_var,
                                   command=self._toggle_merge)
        merge_cb.pack(side="left", padx=(0, 6))
        Tooltip(merge_cb, "Combine all output PDFs into a single merged file with the given filename.")
        self._merge_name_var = ctk.StringVar(value="all.pdf")
        self._merge_entry = ctk.CTkEntry(checks, textvariable=self._merge_name_var, width=100,
                                         state="disabled")
        self._merge_entry.pack(side="left")

        row += 1

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

        self._convert_btn = ctk.CTkButton(
            self, text="Convert", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_convert,
        )
        self._convert_btn.grid(row=row, column=0, padx=16, pady=12, sticky="ew")
        row += 1

        self._progress = ctk.CTkProgressBar(self)
        self._progress.set(0)
        self._progress.grid(row=row, column=0, padx=16, pady=(0, 8), sticky="ew")
        row += 1

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

        start_lbl = ctk.CTkLabel(parent, text="Start Number")
        start_lbl.grid(row=r, column=0, padx=12, pady=6, sticky="w")
        Tooltip(start_lbl, "Counter value for the first output file.\nExample: 5 → MSG_00005.pdf, MSG_00006.pdf, …")
        self._start_var = ctk.StringVar(value="1")
        start_entry = ctk.CTkEntry(parent, textvariable=self._start_var, width=80)
        start_entry.grid(row=r, column=1, padx=(6, 12), pady=6, sticky="w")
        Tooltip(start_entry, "Counter value for the first output file.\nExample: 5 → MSG_00005.pdf, MSG_00006.pdf, …")
        r += 1

        att_lbl = ctk.CTkLabel(parent, text="Max Attachment (MB)")
        att_lbl.grid(row=r, column=0, padx=12, pady=6, sticky="w")
        Tooltip(att_lbl, "Skip attachments larger than this size (in MB).\nSet to 0 for no limit.")
        self._max_att_var = ctk.StringVar(value="10")
        att_entry = ctk.CTkEntry(parent, textvariable=self._max_att_var, width=80)
        att_entry.grid(row=r, column=1, padx=(6, 12), pady=6, sticky="w")
        Tooltip(att_entry, "Skip attachments larger than this size (in MB).\nSet to 0 for no limit.")
        r += 1

        self._bates_var = ctk.BooleanVar(value=False)
        bates_row = ctk.CTkFrame(parent, fg_color="transparent")
        bates_row.grid(row=r, column=0, columnspan=2, padx=12, pady=6, sticky="w")
        bates_cb = ctk.CTkCheckBox(bates_row, text="Bates stamp", variable=self._bates_var)
        bates_cb.pack(side="left", padx=(0, 16))
        Tooltip(bates_cb, "Stamp sequential Bates page numbers on every PDF page.\nUsed for legal document tracking and production sets.")
        pad_lbl = ctk.CTkLabel(bates_row, text="Pad width:")
        pad_lbl.pack(side="left", padx=(0, 6))
        Tooltip(pad_lbl, "Number of digits in Bates numbers.\nExample: pad 6 with prefix 'MSG' → MSG_000001, MSG_000002, …")
        self._bates_pad_var = ctk.StringVar(value="6")
        pad_entry = ctk.CTkEntry(bates_row, textvariable=self._bates_pad_var, width=50)
        pad_entry.pack(side="left")
        Tooltip(pad_entry, "Number of digits in Bates numbers.\nExample: pad 6 with prefix 'MSG' → MSG_000001, MSG_000002, …")
        r += 1

        flags = ctk.CTkFrame(parent, fg_color="transparent")
        flags.grid(row=r, column=0, columnspan=2, padx=12, pady=(4, 10), sticky="w")
        self._nodedup_var = ctk.BooleanVar(value=False)
        dedup_cb = ctk.CTkCheckBox(flags, text="Disable deduplication", variable=self._nodedup_var)
        dedup_cb.pack(side="left", padx=(0, 20))
        Tooltip(dedup_cb, "By default, emails with the same Message-ID are exported only once.\nEnable this to export all copies, including duplicates.")
        self._nonemail_var = ctk.BooleanVar(value=False)
        nonemail_cb = ctk.CTkCheckBox(flags, text="Include non-email items", variable=self._nonemail_var)
        nonemail_cb.pack(side="left")
        Tooltip(nonemail_cb, "Also export calendar entries, contacts, tasks, and other\nnon-email items found in the PST.")
        r += 1

    def _browse_pst(self):
        path = filedialog.askopenfilename(
            title="Select PST file",
            filetypes=[("PST files", "*.pst"), ("All files", "*.*")],
        )
        if path:
            self._pst_var.set(str(Path(path)))
            if not self._out_var.get():
                self._out_var.set(str(Path(path).parent / "output"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output directory")
        if path:
            self._out_var.set(str(Path(path)))

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


    def _poll_log(self):
        try:
            while True:
                msg = _LOG_QUEUE.get_nowait()
                if msg == "__DONE__":
                    self._open_output_dir()
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
    
    def _open_output_dir(self):
        path = self._out_var.get().strip()
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])


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
    
def main():
    _setup_logging()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
