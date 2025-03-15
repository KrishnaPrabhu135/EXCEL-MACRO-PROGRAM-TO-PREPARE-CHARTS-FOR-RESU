# EXCEL-MACRO-PROGRAM-TO-PREPARE-CHARTS-FOR-RESU
Sub ModifyCharts()
Dim cht As ChartObject
For Each cht In Worksheets(1).ChartObjects
cht.Chart.ChartType = xlPie
Next cht Worksheets(1).ChartObjects(1).Activate
ActiveChart.ChartTitle.Text = "Sales Report"
ActiveChart.Legend.Position = xlL
