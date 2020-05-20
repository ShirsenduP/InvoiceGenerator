This simple script takes in a simple excel database of lessons offered to several clients, collects each session by month and student, loads one of two html email templates including payment information or not. Then it generates the formatted emails using another small database of client information and saves it as a Draft ready for a final proofread before sending. 

**SETUP** 
- Change the TEMPLATES variable to include a path to the template emails
- Change the RECORDS variable to be a path to the lesson and contact information database
- Change LESSON_SHEET, CONTACTS_SHEET variables to the names of the sheets within the RECORDS containing lessons and contact information. 
- Change LESSON_INVOICE_HEADERS, CONTACTS_HEADERS to the column headings of lesson information and contact information


**USAGE**
- Change CURRENT_MONTH to whichever month you would like to generate the invoice emails for
- Run ensuring the templates and the excel database is not open