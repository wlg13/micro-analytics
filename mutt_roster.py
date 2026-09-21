#!/usr/bin/env python3
"""mutt display_filter: adds an X-Roster header with the sender's roster info.

Usage in .muttrc:
    set display_filter="/path/to/mutt_roster.py"
    unignore x-roster

Roster CSV: first argument, else $MUTT_ROSTER, else $MA/outputs/merged_roster.csv.
"""
import csv
import os
import sys
from email import message_from_bytes, policy
from email.utils import parseaddr


def roster_path():
    if len(sys.argv) > 1:
        return sys.argv[1]
    if os.environ.get("MUTT_ROSTER"):
        return os.environ["MUTT_ROSTER"]
    return os.path.join(os.environ.get("MA", ""), "outputs", "merged_roster.csv")


def find_student(address):
    with open(roster_path(), newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row["Email"].strip().lower() == address:
                return row
    return None


def header_end(raw):
    ends = [i for i in (raw.find(b"\r\n\r\n"), raw.find(b"\n\n")) if i != -1]
    return min(ends) if ends else -1


def main():
    raw = sys.stdin.buffer.read()
    out = raw
    # Never lose a message: on any failure, pass it through unchanged.
    try:
        end = header_end(raw)
        if end != -1:
            msg = message_from_bytes(raw[:end], policy=policy.default)
            address = parseaddr(str(msg["From"] or ""))[1].lower()
            row = find_student(address) if address else None
            if row:
                first_gen = row["FirstGen"].strip() or "unknown"
                info = " | ".join([
                    f'{row["First Name"]} {row["Last Name"]}',
                    f'ID {row["ID"]}',
                    row["Level"],
                    f"First-gen: {first_gen}",
                    row["Program and Plan"],
                ])
                nl = b"\r\n" if raw[end:end + 2] == b"\r\n" else b"\n"
                out = raw[:end] + nl + f"X-Roster: {info}".encode("utf-8") + raw[end:]
    except Exception:
        pass
    sys.stdout.buffer.write(out)


if __name__ == "__main__":
    main()
