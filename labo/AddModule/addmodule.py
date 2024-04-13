"""
VBAモジュールを既存のExcelに追加して実行する実験

テストコードをリリースから外すための手段を用意する。
"""

import os
import xlwings as xw

moduleDir = os.path.dirname(__file__)
importModulePath = os.path.join(moduleDir, "import.bas")
print(importModulePath)

book: xw.Book = xw.Book("test.xlsm")
try:
    book.api.VBProject.VBComponents.Import(importModulePath)

    funcGetCellValueOriginal = book.macro("GetCellValueOriginal")  # noqa: E501 ここはモジュール名が無くてもいい
    value = funcGetCellValueOriginal(1, 1)
    print(f"original (1,1): {value}")

    sheet: xw.Sheet = book.sheets[0]
    funcGetCellValue = book.macro("mod_ImportCodeForTest_A01002.GetCellValue")  # noqa: E501 モジュール名を省略できない
    value = funcGetCellValue(sheet, 1, 1)
    print(f"(1,1): {value}")
    value = funcGetCellValue(sheet, 2, 3)
    print(f"(2,3): {value}")
    value = funcGetCellValue(sheet, 1, 3)
    print(f"(1,3): {value}")

    funcGetSheet2CellValue = book.macro(
        "mod_ImportCodeForTest_A01002.GetSheet2CellValue"
    )
    value = funcGetSheet2CellValue(1, 1)
    print(f"sheet2 (1,1): {value}")
    value = funcGetSheet2CellValue(2, 2)
    print(f"sheet2 (2,2): {value}")
finally:
    book.close()
