import csv
import argparse
import vobject
from pathlib import Path
from datetime import datetime

# headers from https://support.microsoft.com/en-us/office/create-or-edit-csv-files-to-import-into-outlook-4518d70d-8fe9-46ad-94fa-1494247193c7

CSV_HEADERS = [  # EXACT Microsoft order
    "Title","First Name","Middle Name","Last Name","Suffix","Company","Department",
    "Job Title","Business Street","Business Street 2","Business Street 3",
    "Business City","Business State","Business Postal Code","Business Country/Region",
    "Home Street","Home Street 2","Home Street 3","Home City","Home State",
    "Home Postal Code","Home Country/Region",
    "Other Street","Other Street 2","Other Street 3","Other City","Other State",
    "Other Postal Code","Other Country/Region",
    "Assistant's Phone","Business Fax","Business Phone","Business Phone 2","Callback",
    "Car Phone","Company Main Phone","Home Fax","Home Phone","Home Phone 2","ISDN",
    "Mobile Phone","Other Fax","Other Phone","Pager","Primary Phone","Radio Phone",
    "TTY/TDD Phone","Telex","Account","Anniversary","Assistant's Name",
    "Billing Information","Birthday","Business Address PO Box","Categories","Children",
    "Directory Server","E-mail Address","E-mail Type","E-mail Display Name",
    "E-mail 2 Address","E-mail 2 Type","E-mail 2 Display Name",
    "E-mail 3 Address","E-mail 3 Type","E-mail 3 Display Name",
    "Gender","Government ID Number","Hobby","Home Address PO Box","Initials",
    "Internet Free Busy","Keywords","Language","Location","Manager's Name","Mileage",
    "Notes","Office Location","Organizational ID Number","Other Address PO Box",
    "Priority","Private","Profession","Referred By","Sensitivity","Spouse",
    "User 1","User 2","User 3","User 4","Web Page",
]


def outlook_date(value: str) -> str:
    if not value:
        return ""
    for fmt in ("%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).strftime("%m/%d/%Y")
        except ValueError:
            pass
    return ""


def convert(vcf_path: Path, csv_path: Path) -> None:
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS, quoting=csv.QUOTE_ALL)
        writer.writeheader()

        with open(vcf_path, encoding="utf-8") as vcf:
            for card in vobject.readComponents(vcf):
                row = {h: "" for h in CSV_HEADERS}

                if hasattr(card, "n"):
                    row["First Name"] = card.n.value.given or ""
                    row["Last Name"] = card.n.value.family or ""

                if hasattr(card, "org"):
                    row["Company"] = " ".join(card.org.value)

                if hasattr(card, "title"):
                    row["Job Title"] = card.title.value

                emails = [e.value for e in getattr(card, "email_list", [])]
                if emails:
                    row["E-mail Address"] = emails[0]

                for tel in getattr(card, "tel_list", []):
                    t = [x.upper() for x in tel.params.get("TYPE", [])]
                    if "CELL" in t:
                        row["Mobile Phone"] = tel.value
                    elif "WORK" in t:
                        row["Business Phone"] = tel.value
                    elif "HOME" in t:
                        row["Home Phone"] = tel.value

                if hasattr(card, "bday"):
                    row["Birthday"] = outlook_date(card.bday.value)

                if hasattr(card, "note"):
                    row["Notes"] = card.note.value.replace("\n", " ").strip()

                writer.writerow(row)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("-i", "--input", required=True)
    p.add_argument("-o", "--output", required=True)
    a = p.parse_args()
    convert(Path(a.input), Path(a.output))
    print("CSV generated using Microsoft’s canonical Outlook template.")


if __name__ == "__main__":
    main()
