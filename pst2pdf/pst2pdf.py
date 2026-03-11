#!/usr/bin/env python3
"""pst2pdf — Convert a PST file to one PDF per email for e-discovery."""

from __future__ import annotations

import argparse
import csv
import logging
import sys
from pathlib import Path

from pypdf import PdfWriter

from parser import parse_pst
from renderer import render_email_to_pdf

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="pst2pdf",
        description="Convert a PST file to one PDF per email (e-discovery).",
    )
    parser.add_argument("pst_file", help="Path to the input .pst file")
    parser.add_argument("output_dir", help="Directory to write PDF files into")
    parser.add_argument(
        "--prefix",
        default="MSG",
        help="Filename prefix for sequential naming (default: MSG)",
    )
    parser.add_argument(
        "--start-num",
        type=int,
        default=1,
        metavar="N",
        help="Starting number for sequential naming (default: 1)",
    )
    parser.add_argument(
        "--bates",
        action="store_true",
        help="Burn Bates numbers onto each page (uses --prefix and --start-num)",
    )
    parser.add_argument(
        "--bates-pad",
        type=int,
        default=6,
        metavar="N",
        help="Zero-pad width for Bates page numbers (default: 6 → 000001)",
    )
    parser.add_argument(
        "--manifest",
        action="store_true",
        help="Write a manifest.csv index of all emails to the output directory",
    )
    parser.add_argument(
        "--no-dedup",
        action="store_true",
        help="Disable deduplication — by default emails with duplicate Message-IDs are skipped",
    )
    parser.add_argument(
        "--include-non-email",
        action="store_true",
        help="Include calendar events, contacts, and other non-email items (excluded by default)",
    )
    parser.add_argument(
        "--max-attachment-size",
        type=float,
        default=10.0,
        metavar="MB",
        help="Skip embedding attachments larger than this size in MB (default: 10). Use 0 to disable.",
    )
    parser.add_argument(
        "--attachments",
        choices=["embed", "extract", "none"],
        default="embed",
        help=(
            "How to handle attachments: "
            "'embed' (default) — render as extra pages in each PDF; "
            "'extract' — save original files alongside PDFs as MSG_00001_01_filename.ext; "
            "'none' — ignore attachments entirely"
        ),
    )
    parser.add_argument(
        "--merge",
        metavar="FILENAME",
        help="After converting, merge all PDFs into a single file in the output directory (e.g. --merge all.pdf)",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose/debug logging",
    )
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    pst_path = Path(args.pst_file)
    if not pst_path.is_file():
        logger.error("PST file not found: %s", pst_path)
        sys.exit(1)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    max_attachment_bytes = int(args.max_attachment_size * 1024 * 1024) if args.max_attachment_size > 0 else 0
    # 'none' mode: don't read attachment data at all (-1 = silent skip sentinel)
    # 'extract' or 'embed': read data up to the size limit
    parse_att_bytes = -1 if args.attachments == "none" else max_attachment_bytes

    manifest_rows: list[dict] = []
    output_files: list[Path] = []
    counter = args.start_num
    bates_page_counter = args.start_num
    seen_message_ids: set[str] = set()
    success = 0
    skipped = 0
    deduped = 0

    logger.info("Parsing %s ...", pst_path)

    for email in parse_pst(
        str(pst_path),
        email_only=not args.include_non_email,
        max_attachment_bytes=parse_att_bytes,
    ):
        seq = f"{counter:05d}"
        filename = f"{args.prefix}_{seq}.pdf"
        output_path = output_dir / filename

        # Deduplication by Message-ID
        if not args.no_dedup and email.message_id:
            if email.message_id in seen_message_ids:
                logger.debug("Skipping duplicate Message-ID %s", email.message_id)
                deduped += 1
                continue
            seen_message_ids.add(email.message_id)

        bates_first_page = bates_page_counter

        try:
            bates_next = render_email_to_pdf(
                email,
                str(output_path),
                bates_prefix=args.prefix if args.bates else "",
                bates_start=bates_page_counter,
                bates_pad_width=args.bates_pad,
                embed_attachments=(args.attachments == "embed"),
            )

            # Extract attachments as separate files
            if args.attachments == "extract":
                for i, att in enumerate(email.attachments):
                    if att.is_embedded_msg or not att.data:
                        continue
                    safe_name = "".join(
                        c if c.isalnum() or c in "._- " else "_" for c in att.name
                    ).strip()
                    att_filename = f"{Path(filename).stem}_{i + 1:02d}_{safe_name}"
                    att_path = output_dir / att_filename
                    att_path.write_bytes(att.data)
                    logger.debug("Extracted %s", att_filename)
            if args.bates:
                bates_page_counter = bates_next
                logger.info(
                    "[%s] pages %s–%s  %s",
                    filename,
                    f"{bates_first_page:0{args.bates_pad}d}",
                    f"{bates_next - 1:0{args.bates_pad}d}",
                    email.subject or "(No Subject)",
                )
            else:
                logger.info("[%s] %s", filename, email.subject or "(No Subject)")

            if args.manifest:
                row = {
                    "filename": filename,
                    "date": email.date.isoformat() if email.date else "",
                    "from": email.from_,
                    "to": "; ".join(email.to),
                    "cc": "; ".join(email.cc),
                    "subject": email.subject,
                    "message_id": email.message_id,
                    "folder_path": email.folder_path,
                    "attachments": "; ".join(a.name for a in email.attachments),
                }
                if args.bates:
                    row["bates_begin"] = f"{args.prefix}_{bates_first_page:0{args.bates_pad}d}"
                    row["bates_end"] = f"{args.prefix}_{bates_next - 1:0{args.bates_pad}d}"
                manifest_rows.append(row)

            output_files.append(output_path)
            success += 1
        except Exception as exc:
            logger.warning("Failed to render %s: %s", filename, exc)
            skipped += 1

        counter += 1

    if args.manifest and manifest_rows:
        manifest_path = output_dir / "manifest.csv"
        fieldnames = [
            "filename", "date", "from", "to", "cc",
            "subject", "message_id", "folder_path", "attachments",
        ]
        if args.bates:
            fieldnames += ["bates_begin", "bates_end"]
        with open(manifest_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(manifest_rows)
        logger.info("Manifest written to %s", manifest_path)

    if args.merge and output_files:
        merge_path = output_dir / args.merge
        logger.info("Merging %d PDFs into %s ...", len(output_files), merge_path)
        merger = PdfWriter()
        for pdf in output_files:
            merger.append(str(pdf))
        with open(merge_path, "wb") as f:
            merger.write(f)
        logger.info("Merged PDF written to %s", merge_path)

    total = success + skipped
    if deduped:
        logger.info("Deduplicated: %d duplicate(s) skipped.", deduped)
    logger.info("Done. %d/%d converted. %d skipped.", success, total, skipped)


if __name__ == "__main__":
    main()
