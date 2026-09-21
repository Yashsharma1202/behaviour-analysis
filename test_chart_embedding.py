import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series

wb = Workbook()
ws = wb.active
ws.title = "TestChart"

ws.append(["Trade", "Cum PnL", "Drawdown"])
ws.append([1, 1000, 0])
ws.append([2, 2500, 0])
ws.append([3, 1800, -700])
ws.append([4, 3200, 0])

chart = LineChart()
chart.title = "Cumulative PnL & Drawdown"
chart.style = 13
chart.y_axis.title = "Rupees (₹)"
chart.x_axis.title = "Trade #"

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5)
cats = Reference(ws, min_col=1, min_row=2, max_row=5)

chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "E2")
wb.save("d:/behaviour analysis/test_chart_output.xlsx")
print("Chart created successfully!")
