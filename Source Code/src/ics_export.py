"""
ics_export.py — ICS calendar generation and Outlook COM email helpers.

Generates RFC 5545 .ics files from a solved schedule, with correct
Chicago-timezone → UTC conversion so calendar clients show the right
local times regardless of DST.

Two export modes:
  • "separate" — one .ics per agent (contains only that agent's events)
  • "full"     — one .ics with every agent's events combined

Outlook COM emailing (Windows only) is handled via send_via_outlook().
"""

import os
import uuid
import sys
from datetime import datetime, timezone, timedelta

# zoneinfo is stdlib in 3.9+; tzdata pip package supplies IANA data on Windows
from zoneinfo import ZoneInfo

from schedule_engine_v2 import (
    ALL_SHIFTS, SHIFT_NAMES, SHIFT_TIMES, PHONE_SHIFTS, FD_SHIFTS
)

TIMEZONE = ZoneInfo("UTC")


# ── ICS building helpers ─────────────────────────────────────────────────────

def _fmt_utc(dt_utc):
    """Format a UTC datetime as an ICS DATETIME string (e.g. 20260203T133000Z)."""
    return dt_utc.strftime("%Y%m%dT%H%M%SZ")


def _shift_label(shift_idx):
    """Human-friendly label for the shift, e.g. 'P1 (7:30–10:00 AM)'."""
    sh, sm, eh, em = SHIFT_TIMES[shift_idx]

    def _ampm(h, m):
        suffix = "AM" if h < 12 else "PM"
        display_h = h if h <= 12 else h - 12
        if display_h == 0:
            display_h = 12
        return f"{display_h}:{m:02d} {suffix}"

    return f"{SHIFT_NAMES[shift_idx]}  ({_ampm(sh, sm)}–{_ampm(eh, em)})"


def build_all_shifts_events(generator, year, month_num, edited_assignments=None):
    """
    Build calendar events for ALL shifts in the schedule.
    Each shift becomes a separate event with format: "[Shift Type] - [Agent Name]"
    Phone shifts categorized as Blue, Front Desk as Green.

    Returns a list of event dicts for all shifts.
    """
    if edited_assignments is None:
        edited_assignments = {}

    events = []
    weeks = generator.get_weekly_matrix()

    for week_idx, week_data in enumerate(weeks):
        for day_col in range(5):  # Mon–Fri columns
            date_str = week_data["dates"][day_col]
            if not date_str:
                continue

            for s_idx in ALL_SHIFTS:
                # Check edited overrides first
                key = (week_idx, day_col, s_idx)
                if key in edited_assignments:
                    worker = edited_assignments[key]
                else:
                    worker = week_data["matrix"][s_idx][day_col]

                # Skip empty shifts
                if not worker:
                    continue

                # Parse date_str (e.g. "Feb 03") into a proper date
                try:
                    dt_local_date = datetime.strptime(
                        f"{date_str} {year}", "%b %d %Y"
                    ).date()
                except ValueError:
                    continue

                sh, sm, eh, em = SHIFT_TIMES[s_idx]

                # Build timezone-aware local datetimes
                dt_start_local = datetime(
                    dt_local_date.year, dt_local_date.month, dt_local_date.day,
                    sh, sm, tzinfo=TIMEZONE
                )
                dt_end_local = datetime(
                    dt_local_date.year, dt_local_date.month, dt_local_date.day,
                    eh, em, tzinfo=TIMEZONE
                )

                # Convert to UTC
                dt_start_utc = dt_start_local.astimezone(timezone.utc)
                dt_end_utc = dt_end_local.astimezone(timezone.utc)

                shift_type = "Phone" if s_idx in PHONE_SHIFTS else "Front Desk"
                # Event name format: "[Shift Type] - [Name]"
                summary = f"{shift_type} - {worker}"
                description = f"{worker} – {shift_type} shift ({_shift_label(s_idx)})"

                # Category and color for calendar apps
                # Blue (#3B82F6) for Phone, Green (#10B981) for Front Desk
                if s_idx in PHONE_SHIFTS:
                    category = "Phone"
                    color = "#3B82F6"  # Blue
                else:
                    category = "Front Desk"
                    color = "#10B981"  # Green

                uid = (
                    f"schedulebuild-all-{worker.lower().replace(' ', '')}"
                    f"-{dt_local_date.isoformat()}"
                    f"-shift{s_idx}@helpdesk"
                )

                events.append({
                    "dtstart": dt_start_utc,
                    "dtend": dt_end_utc,
                    "summary": summary,
                    "description": description,
                    "uid": uid,
                    "category": category,
                    "color": color
                })

    return events


def build_agent_events(name, generator, year, month_num, edited_assignments=None):
    """
    Walk through the solved schedule and collect calendar events for one agent.

    Returns a list of dicts:
      [{"dtstart": <utc>, "dtend": <utc>, "summary": str, "description": str,
        "uid": str}, ...]
    """
    if edited_assignments is None:
        edited_assignments = {}

    events = []
    weeks = generator.get_weekly_matrix()

    for week_idx, week_data in enumerate(weeks):
        for day_col in range(5):  # Mon–Fri columns
            date_str = week_data["dates"][day_col]
            if not date_str:
                continue

            for s_idx in ALL_SHIFTS:
                # Check edited overrides first
                key = (week_idx, day_col, s_idx)
                if key in edited_assignments:
                    worker = edited_assignments[key]
                else:
                    worker = week_data["matrix"][s_idx][day_col]

                if worker != name:
                    continue

                # Parse date_str (e.g. "Feb 03") into a proper date
                try:
                    dt_local_date = datetime.strptime(
                        f"{date_str} {year}", "%b %d %Y"
                    ).date()
                except ValueError:
                    continue

                sh, sm, eh, em = SHIFT_TIMES[s_idx]

                # Build timezone-aware local datetimes
                dt_start_local = datetime(
                    dt_local_date.year, dt_local_date.month, dt_local_date.day,
                    sh, sm, tzinfo=TIMEZONE
                )
                dt_end_local = datetime(
                    dt_local_date.year, dt_local_date.month, dt_local_date.day,
                    eh, em, tzinfo=TIMEZONE
                )

                # Convert to UTC
                dt_start_utc = dt_start_local.astimezone(timezone.utc)
                dt_end_utc = dt_end_local.astimezone(timezone.utc)

                shift_type = "Phone" if s_idx in PHONE_SHIFTS else "Front Desk"
                # NEW: Event name format: "[Shift Type] - [Name]"
                summary = f"{shift_type} - {name}"
                description = f"{name} – {shift_type} shift ({_shift_label(s_idx)})"

                # Category and color for calendar apps
                # Blue (#3B82F6) for Phone, Green (#10B981) for Front Desk
                if s_idx in PHONE_SHIFTS:
                    category = "Phone"
                    color = "#3B82F6"  # Blue
                else:
                    category = "Front Desk"
                    color = "#10B981"  # Green

                uid = (
                    f"schedulebuild-{name.lower().replace(' ', '')}"
                    f"-{dt_local_date.isoformat()}"
                    f"-shift{s_idx}@helpdesk"
                )

                events.append({
                    "dtstart": dt_start_utc,
                    "dtend": dt_end_utc,
                    "summary": summary,
                    "description": description,
                    "uid": uid,
                    "category": category,
                    "color": color
                })

    return events


def _build_vcalendar(events, cal_name="Helpdesk Schedule", include_reminder=False):
    """
    Turn a list of event dicts into a full RFC 5545 VCALENDAR string.
    """
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ScheduleBuilder//HelpDesk//EN",
        f"X-WR-CALNAME:{cal_name}",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]

    for ev in events:
        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{ev['uid']}",
            f"DTSTART:{_fmt_utc(ev['dtstart'])}",
            f"DTEND:{_fmt_utc(ev['dtend'])}",
            f"SUMMARY:{ev['summary']}",
            f"DESCRIPTION:{ev['description']}",
            "STATUS:CONFIRMED",
            "TRANSP:OPAQUE",
        ])

        # Add category for organization
        if "category" in ev:
            lines.append(f"CATEGORIES:{ev['category']}")

        # Add COLOR property (RFC 7986) for visual color coding
        if "color" in ev:
            lines.append(f"COLOR:{ev['color']}")

        if include_reminder:
            lines.extend([
                "BEGIN:VALARM",
                "TRIGGER:-PT15M",
                "ACTION:DISPLAY",
                f"DESCRIPTION:Reminder: {ev['summary']}",
                "END:VALARM"
            ])

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


# ── Public API ───────────────────────────────────────────────────────────────

def write_ics_single_agent(agent_name, events, filepath):
    """Write one .ics file containing only this agent's events."""
    cal_text = _build_vcalendar(events, cal_name=f"{agent_name} – Helpdesk Schedule", include_reminder=True)
    with open(filepath, "w", newline="") as f:
        f.write(cal_text)


def write_ics_full_schedule(all_shifts_events, filepath):
    """
    Write a single .ics with all shifts as separate events.
    Each event is named "[Shift Type] - [Agent Name]" with color categories.
    all_shifts_events: list of event dicts (from build_all_shifts_events)
    """
    # Sort by start time for readability
    all_shifts_events.sort(key=lambda e: e["dtstart"])
    cal_text = _build_vcalendar(all_shifts_events, cal_name="Full Helpdesk Schedule")
    with open(filepath, "w", newline="") as f:
        f.write(cal_text)


def generate_all_ics(generator, employees, edited_assignments,
                     year, month_num, output_folder, mode="separate"):
    """
    High-level orchestrator.

    mode:
      "separate" — creates <output_folder>/<Name>.ics for each agent
      "full"     — creates <output_folder>/Full_Schedule.ics

    Returns list of created file paths.
    """
    os.makedirs(output_folder, exist_ok=True)
    created = []

    if mode == "separate":
        # Build individual agent calendars
        for emp in employees:
            name = emp["name"]
            events = build_agent_events(name, generator, year, month_num, edited_assignments)
            if not events:
                continue
            safe_name = name.replace(" ", "_").replace("/", "_")
            path = os.path.join(output_folder, f"{safe_name}.ics")
            write_ics_single_agent(name, events, path)
            created.append(path)

    elif mode == "full":
        # Build full schedule with all shifts as separate color-coded events
        all_shifts = build_all_shifts_events(generator, year, month_num, edited_assignments)
        path = os.path.join(output_folder, "Full_Schedule.ics")
        write_ics_full_schedule(all_shifts, path)
        created.append(path)

    else:
        raise ValueError(f"Unknown mode: {mode!r}. Use 'separate' or 'full'.")

    return created


# ── Outlook COM Email ────────────────────────────────────────────────────────

def is_outlook_available():
    """
    Returns True if pywin32 + Outlook COM dispatch is importable.
    Always False on non-Windows platforms.
    """
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def send_via_outlook(recipients, subject, body_html, attachment_paths=None):
    """
    Open a *draft* Outlook email (MailItem.Display) so the user can review
    before sending. Raises RuntimeError if Outlook cannot be accessed.

    recipients: list of (name, email) tuples
    attachment_paths: list of absolute file paths to attach
    """
    if not is_outlook_available():
        raise RuntimeError(
            "Outlook COM automation is not available on this system.\n"
            "Ensure pywin32 is installed and Outlook is configured."
        )

    import win32com.client

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError(
            f"Could not launch Outlook:\n{e}\n\n"
            "Make sure Outlook is installed and running."
        )

    mail = outlook.CreateItem(0)  # olMailItem = 0
    mail.Subject = subject
    mail.HTMLBody = body_html

    # Semicolon-separated recipients
    to_list = "; ".join(email for _, email in recipients if email)
    mail.To = to_list

    if attachment_paths:
        for path in attachment_paths:
            abs_path = os.path.abspath(path)
            if os.path.isfile(abs_path):
                mail.Attachments.Add(abs_path)

    mail.Display()  # Opens the draft for review — does NOT auto-send
    return True


def build_email_body(month_name, year, employees):
    """
    Build a simple HTML email body for the schedule distribution.
    """
    html = f"""
    <html>
    <body style="font-family: Calibri, Arial, sans-serif; color: #333;">
    <h2 style="color: #3B82F6;">Helpdesk Schedule — {month_name} {year}</h2>
    <p>Hi team,</p>
    <p>Please find your schedule files attached:</p>
    <ul>
        <li><strong>Your individual calendar</strong> (<code>.ics</code>) — Double-click to import into Outlook, Google Calendar, etc.</li>
        <li><strong>Full schedule spreadsheet</strong> (<code>.xlsx</code>) — View the complete team schedule with the fairness audit</li>
    </ul>
    <p>If you have questions or need to swap a shift, please reach out.</p>
    </body>
    </html>
    """
    return html.strip()
