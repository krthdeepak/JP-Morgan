#######################################################################
#
# An example of creating Excel Stock charts with Python and XlsxWriter.
#
# Copyright 2013-2017, John McNamara, jmcnamara@cpan.org
#
from datetime import datetime
import xlsxwriter
from excel_styles import ExcelStyles

workbook = xlsxwriter.Workbook('international_market.xlsx')
worksheet = workbook.add_worksheet('Trade Rate')
worksheet.set_tab_color('green')

# Add a filter to the worksheet.
worksheet.autofilter('G1:H1')

bold = workbook.add_format({'bold': 1,'align': 'center','bg_color':'yellow'})
center_align = workbook.add_format({'align': 'center'})

date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})

chart = workbook.add_chart({'type': 'stock'})

# Add the worksheet data that the charts will refer to.
headings = ['Entity', 'Buy/Sell', 'AgreedFx', 'Currency', 'Units', 'Price per unit','Incoming(USD)', 'Outgoing(USD)']
data = {('foo','B',0.50,'SGP','01-01-2016','31 Jan 2016',200,100.25),
        ('bar','S',0.22,'AED','05-01-2016','10 Jan 2016',450,150.50),
        ('foo1','B',0.50,'SGP','10-01-2016','20 Jan 2016',300,125.25),
        ('bar1','S',0.22,'AED','25-01-2016','29 Jan 2016',550,250.50),
        ('sar','S',0.22,'AED','15-01-2016','20 Jan 2016',400,100.50),
        ('var','S',0.25,'AED','04-01-2016','19 Jan 2016',300,120.50)}

worksheet.write_row('A1', headings, bold)
row = 1

for i in data:

    worksheet.write(row, 0, i[0])
    worksheet.write(row, 1, i[1])
    worksheet.write(row, 2, i[2])
    worksheet.write(row, 3, i[3])
    worksheet.write(row, 4, i[6])
    worksheet.write(row, 5, i[7])
    if i[1] == 'S':
        worksheet.write(row, 6, float(i[7]) * int(i[6])* float(i[2]))
    else:
        worksheet.write(row, 6, 0)
    if i[1] == 'B':
        worksheet.write(row, 7, float(i[7]) * int(i[6])* float(i[2]))
    else:
        worksheet.write(row, 7, 0)            
    row += 1
    
worksheet.set_column('A:I', 15)

workbook.close()
