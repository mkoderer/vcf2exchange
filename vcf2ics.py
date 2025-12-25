import argparse
import hashlib
import sys
from datetime import date, datetime, timedelta, time
from pathlib import Path

import vobject
from icalendar import (
    Calendar,
    Event,
    Alarm,
    Timezone,
    TimezoneStandard,
    TimezoneDaylight,
)

TZID = "Europe/Berlin"


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


def berlin_timezone():
    tz = Timezone()
    tz.add("tzid", TZID)

    standard = TimezoneStandard()
    standard.add("dtstart", datetime(1970, 10, 25, 3, 0, 0))
    standard.add("tzoffsetfrom", timedelta(hours=2))
    standard.add("tzoffsetto", timedelta(hours=1))
    standard.add("tzname", "CET")

    daylight = TimezoneDaylight()
    daylight.add("dtstart", datetime(1970, 3, 29, 2, 0, 0))
    daylight.add("tzoffsetfrom", timedelta(hours=1))
    daylight.add("tzoffsetto", timedelta(hours=2))
    daylight.add("tzname", "CEST")

    tz.add_component(standard)
    tz.add_component(daylight)
    return tz


def main():
    parser = argparse.ArgumentParser(
        description="Create an Outlook-compatible birthday calendar (ICS) from VCF files."
    )
    parser.add_argument("-i", "--input", nargs="+", required=True)
    parser.add_argument("-o", "--output", help="Output ICS file (default: STDOUT)")
    parser.add_argument(
        "--reminder-time",
        default="09:00",
        help="Reminder time HH:MM local time (default: 09:00)",
    )

    args = parser.parse_args()
    reminder_hour, reminder_minute = map(int, args.reminder_time.split(":"))

    cal = Calendar()
    cal.add("prodid", "-//VCF Birthday Calendar//EN")
    cal.add("version", "2.0")

    # Add explicit Berlin timezone
    cal.add_component(berlin_timezone())

    today = date.today()

    for vcf_file in args.input:
        with open(vcf_file, encoding="utf-8") as f:
            for card in vobject.readComponents(f):
                if not hasattr(card, "bday"):
                    continue

                birthday = parse_birthday(card.bday.value)
                if not birthday:
                    continue

                name = " ".join(
                    filter(
                        None,
                        [
                            getattr(card.n.value, "given", ""),
                            getattr(card.n.value, "family", ""),
                        ],
                    )
                ) or "Unknown"

                event_date = birthday.replace(year=today.year)

                event = Event()
                event.add("uid", make_uid(name, birthday))
                event.add("summary", f"Geburtstag: {name}")
                event.add("description", f"Gebrutstag von {name} am {birthday.isoformat()}")

                # All-day event (DATE, not DATE-TIME)
                event.add("dtstart", event_date)
                event.add("dtend", event_date + timedelta(days=1))

                # Yearly recurrence
                event.add(
                "rrule",
                    {
                        "freq": "yearly",
                        "interval": 1,
                        "bymonth": birthday.month,
                        "bymonthday": birthday.day,
                    },
                )

                # Alarm at local time with TZID
                alarm = Alarm()
                alarm.add("action", "DISPLAY")

                trigger_dt = datetime.combine(
                    event_date,
                    time(reminder_hour, reminder_minute),
                )
                alarm.add("trigger", trigger_dt)
                alarm["trigger"].params["TZID"] = TZID

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
