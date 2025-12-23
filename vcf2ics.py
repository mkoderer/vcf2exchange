import argparse
import hashlib
import sys
from datetime import date, datetime, timedelta, time
from pathlib import Path

import pytz
import vobject
from icalendar import Calendar, Event, Alarm


def parse_birthday(value: str):
    for fmt in ("%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def make_uid(name: str, birthday: date) -> str:
    base = f"{name}-{birthday.isoformat()}"
    digest = hashlib.sha1(base.encode("utf-8")).hexdigest()
    return f"birthday-{digest}@local"


def main():
    parser = argparse.ArgumentParser(
        description="Create an Outlook-compatible birthday calendar (ICS) from VCF files."
    )
    parser.add_argument("-i", "--input", nargs="+", required=True)
    parser.add_argument("-o", "--output", help="Output ICS file (default: STDOUT)")
    parser.add_argument(
        "--timezone",
        default="UTC",
        help="IANA timezone (e.g. Europe/Berlin)"
    )
    parser.add_argument(
        "--reminder-time",
        default="09:00",
        help="Reminder time HH:MM (default: 09:00)"
    )

    args = parser.parse_args()

    tz = pytz.timezone(args.timezone)
    reminder_hour, reminder_minute = map(int, args.reminder_time.split(":"))

    cal = Calendar()
    cal.add("prodid", "-//VCF Birthday Calendar//EN")
    cal.add("version", "2.0")

    today = date.today()

    for vcf_file in args.input:
        with open(vcf_file, encoding="utf-8") as f:
            for card in vobject.readComponents(f):
                if not hasattr(card, "bday"):
                    continue

                birthday = parse_birthday(card.bday.value)
                if not birthday:
                    continue

                # Build display name
                name = " ".join(
                    filter(None, [
                        getattr(card.n.value, "given", ""),
                        getattr(card.n.value, "family", "")
                    ])
                ) or "Unknown"

                # Event date in current year
                event_date = birthday.replace(year=today.year)

                event = Event()
                event.add("uid", make_uid(name, birthday))
                event.add("summary", name)
                event.add("description", f"Birthday: {birthday.isoformat()}")
                event.add("dtstart", event_date)
                event.add("dtend", event_date + timedelta(days=1))
                event.add("rrule", {"freq": "yearly"})

                # Reminder at specific local time
                alarm = Alarm()
                alarm.add("action", "DISPLAY")

                trigger_dt = tz.localize(
                    datetime.combine(event_date, time(reminder_hour, reminder_minute))
                )
                alarm.add("trigger", trigger_dt)

                event.add_component(alarm)
                cal.add_component(event)

    ics_bytes = cal.to_ical()

    if args.output:
        Path(args.output).write_bytes(ics_bytes)
        print(f"Birthday calendar written to {args.output}", file=sys.stderr)
    else:
        sys.stdout.buffer.write(ics_bytes)


if __name__ == "__main__":
    main()
