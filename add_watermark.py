from openpyxl import load_workbook

wb = load_workbook(r'C:\Users\PrashantKumar\csv2excel\CareFirst_Governance_Report.xlsx')

header_text = '&"Calibri,Bold"&36&KA6A6A6CONFIDENTIAL \u2014 PROPRIETARY TO AHEAD'
footer_left  = '&"Calibri,Regular"&9&K808080Confidential \u2014 Proprietary to AHEAD'
footer_right = '&"Calibri,Regular"&9&K808080Page &P of &N'

for ws in wb.worksheets:
    ws.oddHeader.center.text = header_text
    ws.oddFooter.left.text   = footer_left
    ws.oddFooter.right.text  = footer_right

wb.save(r'C:\Users\PrashantKumar\csv2excel\CareFirst_Governance_Report.xlsx')
print('Watermark added to tabs:', wb.sheetnames)
