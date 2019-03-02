import xlsxwriter
import random

random_data = [23,24,56,23]

data_start_loc = [0, 0]
data_end_loc = [data_start_loc[0] + len(random_data)-1, 0]

workbook = xlsxwriter.Workbook('E:\\abc.xlsx')


chart = workbook.add_chart({'type': 'column'})
chart.set_y_axis({'name': 'Values'})
chart.set_x_axis({'name': 'Sequential order'})
chart.set_title({'name': 'Graph in Excel through Python'})


worksheet = workbook.add_worksheet()

worksheet.write_column(*data_start_loc, data=random_data)

chart.add_series({
    'values': [worksheet.name] + data_start_loc + data_end_loc,
    'name': "Random data",
    'points': [
        {'fill': {'color': 'red'}},
        {'fill': {'color': 'green'}},
        {'fill': {'color': 'blue'}},
        {'fill': {'color': 'gray'}},
    ],
    'border': {'color': 'black'}
})

chart.set_plotarea({
    'fill':   {'color': 'yellow'}
    # 'pattern': {
    #     'pattern': 'percent_25',
    #     'fg_color': 'red',
    #     'bg_color': 'yellow',
    # }
})

worksheet.insert_chart('B1', chart)

workbook.close()
