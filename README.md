# office-ps-automation
Powershell script to automate simple recurrent tasks with excel and outlook
This script is intended as a template for more complex scenarios.

## Requirements
- Modify variables to suite your needs (Paths, emails, etc.).
- Itextsharp 5.5.13. Modify the route of the dll in line 118. Itextsharp is used to check the content of generated pdfs.

## Behaviour

- Modify excel.
- Generate pdf of modified sheet.
- Create mail attaching generated pdf.
- Check if keywords are present in generated pdf.
- Send if correct. Else send mail to administrator with error.
