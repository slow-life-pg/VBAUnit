Attribute VB_Name = "mod_ImportCodeForTest_A01002"

Option Explicit

Public Function GetCellValue(ByRef sheet As Worksheet, ByVal row As Integer, ByVal col As Integer)
    GetCellValue = sheet.Cells(row, col).Value
End Function

Public Function GetSheet2CellValue(ByVal row As Integer, ByVal col As Integer)
    GetSheet2CellValue = Sheet2.Cells(row, col).Value
End Function
