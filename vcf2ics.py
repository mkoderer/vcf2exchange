import argparse
import hashlib
import sys
from datetime import datetime, timedelta, time, date
from pathlib import Path

import pytz
import vobject
from ics import Calendar, Event, DisplayAlarm
from ics.grammar.parse import ContentLine


def parse_birthday(value: str):
    for fmt in ("%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def make_uid(name: str, date) -> str:
    base = f"{name}-{date.isoformat()}"
    digest = hashlib.sha1(base.encode("utf-8")).hexdigest()
    return f"birthday-{digest}@local"


def main():
    parser = argparse.ArgumentParser(
        description="Create an Outlook-compatible birthday calendar (all-day events with timed reminders)."
    )
    parser.add_argument("-i", "--input", nargs="+", required=True)
    parser.add_argument("-o", "--output", help="Output ICS file (default: STDOUT)")
    parser.add_argument(
        "--timezone",
        default="UTC",
        help="IANA timezone name (e.g. Europe/Berlin)"
    )
    parser.add_argument(
        "--reminder-time",
        default="09:00",
        help="Local reminder time HH:MM (default: 09:00)"
    )

    args = parser.parse_args()

    tz = pytz.timezone(args.timezone)
    hour, minute = map(int, args.reminder_time.split(":"))
    reminder_clock = time(hour, minute)

    calendar = Calendar()
  
    for vcf_file in args.input:
        with open(vcf_file, encoding="utf-8") as f:
            for card in vobject.readComponents(f):
                if not hasattr(card, "bday"):
                    continue

                birthday = parse_birthday(card.bday.value)
                # Replace with the current year
                today = date.today()
                birthday_this_year = birthday.replace(year=today.year)
                if not birthday_this_year:
                    continue

                name = " ".join(
                    filter(None, [
                        getattr(card.n.value, "given", ""),
                        getattr(card.n.value, "family", "")
                    ])
                ) or "Unknown"

                # All-day event
                event = Event()
                event.name = f"Birthday: {name}"
                event.begin = birthday_this_year
                event.make_all_day()
                event.uid = make_uid(name, birthday)
                event.extra.append(ContentLine(name="RRULE", value="FREQ=YEARLY"))
                event.description = f"Birthday: {birthday.isoformat()}"  # YYYY-MM-DD
                
                # Absolute alarm at local time
                alarm_dt = tz.localize(
                    datetime.combine(birthday_this_year, reminder_clock)
                )
                minutes = hour * 60 + minute
                event.alarms.append(
                    DisplayAlarm(trigger=timedelta(minutes=minutes))
                )

                calendar.events.add(event)

    output = calendar.serialize()

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Birthday calendar written to {args.output}", file=sys.stderr)
    else:
        sys.stdout.write(output)


if __name__ == "__main__":
    main()
