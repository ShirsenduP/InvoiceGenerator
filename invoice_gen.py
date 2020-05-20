import sys
import numpy as np
import pandas as pd
import argparse
import os
from tabulate import tabulate

CURRENT_MONTH = "May"  # TODO: Convert into command line script where month is argument

TEMPLATES = {
    "payment": "c:/Users/spodd/Google Drive/Students/EVT Timesheets/template_payment.txt",
    "nopayment": "c:/Users/spodd/Google Drive/tudents/EVT Timesheets/template_nopayment.txt",
}

RECORDS = "c:/Users/spodd/Onedrive - University College London/Administration.xlsx"
LESSON_SHEET = "ST_Detailed"
LESSON_INVOICE_HEADERS = {
    "Student": "Student",
    "Start Time": "Start time",
    "Length": "Hrs",
    "Rate": "Rate",
    "Paid": "Settled",
    "Unpaid": "Outstanding",
}
CONTACTS_SHEET = "ST_Info"
CONTACTS_HEADERS = {
    "Student": "Student",
    "Contact": "Contact",
    "Email": "Email",
}


def draft(text, subject, recipient):
    import win32com.client as win32

    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Save()


if __name__ == "__main__":

    # Load Excel file with all information
    try:
        df = pd.ExcelFile(RECORDS)
    except PermissionError as e:
        print(e)
        sys.exit(
            "PermissionError: Make sure the file is closed before running the script."
        )
    else:
        details = df.parse(CONTACTS_SHEET, index_col=CONTACTS_HEADERS["Student"])
        lesson_records = df.parse(LESSON_SHEET)

    # Extract information for current month only
    current_month = lesson_records.groupby("Month").get_group(CURRENT_MONTH)
    lessons_by_student = current_month.groupby(LESSON_INVOICE_HEADERS["Student"])

    # Generate draft for each student
    for name, lessons in lessons_by_student:

        # Prepare invoice
        invoice = np.round(lessons[LESSON_INVOICE_HEADERS.values()], decimals=2,)

        # Change index to list how many lessons there were for each student
        invoice.index = np.arange(1, len(invoice) + 1)

        # Convert to HTML
        pretty_invoice = tabulate(invoice, headers="keys", tablefmt="html")

        # Calculate total due
        total_due = sum(invoice[LESSON_INVOICE_HEADERS["Unpaid"]])

        # Open template with/without payment information
        template = TEMPLATES["payment"] if total_due > 0 else TEMPLATES["nopayment"]
        with open(template, "r") as f:
            msg_template = f.read()

        # Get contact information: Name and Email Address
        info = details.loc[name]
        parent_name = info[0].split()[0]  # Get first name only
        contact_email = info[CONTACTS_HEADERS["Email"]]

        # Format the email body
        message = msg_template.format(
            name=parent_name, month=CURRENT_MONTH, table=pretty_invoice, due=total_due
        )

        # Draft emails
        print(f"Compiling email for {name}")
        draft(message, f"EVT Invoice for {CURRENT_MONTH}", contact_email)

    print("Completed")
