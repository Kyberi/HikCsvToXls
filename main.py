import csv
import pprint
import xlsxwriter

start_row = 1
start_col = 1

workbook = xlsxwriter.Workbook('result.xlsx')
# worksheet = workbook.add_worksheet()

# Column id in csv
num_id = 0
num_name = 1  # Name employees
num_date = 2  # Date format yyyy-mmm-dd
num_time = 3  # Time format hh:mm
k = True
with open('full.csv', mode='r', encoding='utf8') as f:
    reader = csv.reader(f, delimiter=',')
    d = {}
    # person_in = 0
    # person_out = 0

    for row in reader:
        if not row[num_id] in d:
            d[row[num_id]] = {}
            d[row[num_id]]['id'] = row[num_id]
            d[row[num_id]]['name'] = row[num_name]
            d[row[num_id]]['schedule'] = {}
        # Split date
        date = row[num_date].split("-")
        year = date[2]
        month = date[1]
        day = date[0]
        # Parsing years
        if year not in d[row[num_id]]['schedule']:
            d[row[num_id]]['schedule'] = {}
            d[row[num_id]]['schedule'][year] = {}
        # Parsing months
        if month not in d[row[num_id]]['schedule'][year]:
            d[row[num_id]]['schedule'][year][month] = {}
        # Parsing days
        if day not in d[row[num_id]]['schedule'][year][month]:
            d[row[num_id]]['schedule'][year][month][day] = {}
            k = True
        # Split hour
        time = row[num_time].split(":")
        hours = time[0]
        minutes = time[1]

        if k:
            h_in = hours
            m_in = minutes
            k = False
        h_out = hours
        m_out = minutes

        d[row[num_id]]['schedule'][year][month][day]['in'] = h_in + ":" + m_in
        d[row[num_id]]['schedule'][year][month][day]['out'] = h_out + ":" + m_out

# print(d)
pprint.pprint(d)


for pid in d.keys():
    worksheet = workbook.add_worksheet(d[pid]['name'])
    # print(d[pid]['name'])
    # worksheet.write(start_row + row, start_col, pid)

    for year in d[pid]['schedule'].keys():
        col = 0
        for month in d[pid]['schedule'][year].keys():
            worksheet.write(start_row, start_col + (col * 3), month)
            worksheet.write(start_row, start_col + (col * 3) + 1, 'Вхід')
            worksheet.write(start_row, start_col + (col * 3) + 2, 'Вихід')
            row = 0
            for day in d[pid]['schedule'][year][month].keys():
                worksheet.write(start_row + 1 + row, start_col + (col * 3), int(day))
                worksheet.write(start_row + 1 + row, start_col + (col * 3) + 1,
                                d[pid]['schedule'][year][month][day]['in'])
                worksheet.write(start_row + 1 + row, start_col + (col * 3) + 2,
                                d[pid]['schedule'][year][month][day]['out'])

                row += 1
            col += 1
        pass


workbook.close()
