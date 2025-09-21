"""
ブックを既存のブックにアドインとして連結して実行する実験

テストコードをテストツールから操作するためのBridgeを実現する。
"""

import os
import traceback as tb
import xlwings as xw

moduleDir = os.path.dirname(__file__)
baseBookPath = os.path.join(moduleDir, "bridge.xlsm")
addInBookPath = os.path.join(moduleDir, "addition.xlsm")
print(addInBookPath)

app = xw.App(visible=False)
try:
    baseBook = app.books.open(baseBookPath, update_links=True, ignore_read_only_recommended=True)

    funcGetBaseCellValue = baseBook.macro("GetBaseCellValue")
    value = funcGetBaseCellValue(1, 1)
    print(f"Base self (1,1): {value}")

    # baseBook.api.VBProject.References.AddFromFile(addInBookPath)
    app.books.open(addInBookPath)

    addInBook: xw.Book = app.books["addition.xlsm"]

    baseShee: xw.Sheet = baseBook.sheets["BaseSheet"]
    addINSheet: xw.Sheet = addInBook.sheets["Sheet1"]

    print("validation:")

    funcGetBaseSheetCellValue = baseBook.macro("GetBaseSheetCellValue")
    value = funcGetBaseSheetCellValue(baseShee, 2, 2)
    print(f"Base self (2,2): {value}")
    value = funcGetBaseSheetCellValue(addINSheet, 1, 1)
    print(f"AddIn from Base (1,1): {value}")

    funcGetAddInCellValue = baseBook.macro("addition.xlsm!GetCellValue")
    funcGetAddInCellValue = addInBook.macro("GetCellValue")
    value = funcGetAddInCellValue(1, 1)
    print(f"AddIn self (1,1): {value}")

    funcGetAddInSheetCellValue = baseBook.macro("addition.xlsm!GetSheetCellValue")
    funcGetAddInSheetCellValue = addInBook.macro("GetSheetCellValue")
    value = funcGetAddInSheetCellValue(baseShee, 1, 1)
    print(f"Base from AddIn (1,1): {value}")

except Exception as e:
    print("mainflow error:")
    print(tb.print_exception(e))
finally:
    try:
        # addInRef = None
        # addInName: str = app.books["addition.xlsm"].api.VBProject.name
        # for ref in app.books["bridge.xlsm"].api.VBProject.References:
        #     if ref.name == addInName:
        #         addInRef = ref
        #         break
        # if addInRef:
        #     print(f"ref: {addInRef.name}")
        #     app.books["bridge.xlsm"].api.VBProject.References.Remove(addInRef)
        # else:
        #     print("ref is not found!!!")

        print("quit")
        app.quit()
    except Exception as e:
        print("finally section error:")
        print(tb.print_exception(e))
        app.kill()
