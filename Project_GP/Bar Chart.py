import xlsxwriter
import pandas as pd
from collections import Counter
import math
import numpy as np

workbook = xlsxwriter.Workbook('Bar_chart.xlsx')
worksheet = workbook.add_worksheet()


f_n = 'data2_moderate'
out_path = f_n + '_out.xlsx'
df = pd.read_excel(f_n + '.xlsx', sheet_name=2, skiprows=0)

def count_range(li, min, max):
	ctr = 0
	for x in li:
		if min <= x <= max:
			ctr += 1
	return ctr


headings = list(df.columns[0:4])
headings.extend(['Range', 'Count Loss', 'Unique Zone', 'Count Zone'])

print(headings)

nodeB = df['NodeB'].tolist()
ippm_loss = df['IPPM loss(%)'].tolist()
district = df['District'].tolist()
tp_zone = df['TP Zone'].tolist()


district2=['NA' if x is np.nan else x for x in district]

#print(district2)





range_ippm_loss = ['0-5', '6-10','11-15', '16-20','21-25', '26-30','31-35', '36-40','41-45', '46-50','51-55', '56-60','61-65', '66-70','71-75', '76-80',]
count_ippm_loss = []


for i in range(0, 80, 5):
    count_ippm_loss.append(count_range(ippm_loss, i, i+5))


# Create a new Chart object.
chart = workbook.add_chart({'type': 'column'})

chart2 = workbook.add_chart({'type': 'column'})

# Add a format for the headings.
bold = workbook.add_format({'bold': 1})

#print('range' + range_ippm_loss)
# Write some data to add to plot on the chart.

#data = [range_ippm_loss, count_ippm_loss ]

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', nodeB)
worksheet.write_column('B2', ippm_loss)
worksheet.write_column('C2', district2)
worksheet.write_column('D2', tp_zone)
worksheet.write_column('E2', range_ippm_loss)
worksheet.write_column('F2', count_ippm_loss)



# Configure the chart. In simplest case we add one or more data series.
chart.add_series({
	'categories': '=Sheet1!$E$2:$E$17',
    'values': '=Sheet1!$F$2:$F$17',
    'data_labels': {'value': True},
    'fill':   {'color': 'green'},
    'border': {'color': 'black'}
})

chart.set_size({'width': 700, 'height': 400})

# Add a chart title and some axis labels.
chart.set_title ({'name': 'Site count of IPPM loss(%) range'})
chart.set_x_axis({'name': 'IPPM Loss(%) in range'})
chart.set_y_axis({'name': 'Count of IPPM Loss(%)'})


# Insert the chart into the worksheet.
worksheet.insert_chart('I20', chart)


#---------------------------------------for TP Zone---------------------------------
zz = []
zones = list(Counter(tp_zone).keys())
zones_no = list(Counter(tp_zone).values())

for i in range(len(zones)):
    zz.append([zones[i], zones_no[i]])


zz = sorted(zz, key = lambda x: int(x[1]))

zones.clear()
zones_no.clear()

for i in range(len(zz)):
    zones.append(zz[i][0])
    zones_no.append(zz[i][1])





worksheet.write_column('G2', zones)
worksheet.write_column('H2', zones_no)

chart2.add_series({
	'categories': '=Sheet1!$G$2:$G$' + str(len(zones)+1),
    'values': '=Sheet1!$H$2:$H$'+ str(len(zones)+1),

    'data_labels': {'value': True},
    'fill':   {'color': 'blue'},
    'border': {'color': 'black'}
})


# Add a chart title and some axis labels.
chart2.set_title ({'name': 'TP Zone wise count'})
chart2.set_x_axis({'name': 'Unique Zone'})
chart2.set_y_axis({'name': 'Count of Zone'})
chart2.set_legend({'none': True})

# Insert the chart into the worksheet.
worksheet.insert_chart('I2', chart2)




workbook.close()

#input("tap any")
