# VBA-challenge

Following code to find the last row copied from https://www.thespreadsheetguru.com/last-row-column-vba/
Dim LastRow As Long
LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
