"""
ブックのVBAコードを吸い出す実験

テスト実行時のスナップショットを保存するために

※ThisWorkbookオブジェクトとSheetオブジェクトを区別できなかった。
"""

import os
import pathlib
import traceback as tb
import xlwings as xw

moduleDir = os.path.dirname(__file__)
exportDir = os.path.join(moduleDir, "vba_files")
bookPath = os.path.join(moduleDir, "source.xlsm")

for existingFile in pathlib.Path(exportDir).iterdir():
    os.remove(existingFile)

book = xw.Book(bookPath)
try:
    for module in book.api.VBProject.VBComponents:
        print(module.name)
        ext = ".txt"
        if module.Type == 1:
            print("  is module")
            ext = ".bas"
        elif module.Type == 2:
            print("  is class")
            ext = ".cls"
        elif module.Type == 3:
            print("  is form")
            ext = ".frm"
        elif module.Type == 100:
            print("  is book or sheet")
            ext = ".bas"
        else:
            print(f"  is type {module.Type}")
        
        vbafilePath = os.path.join(exportDir, f"{module.name}{ext}")
        module.Export(vbafilePath)

except Exception as e:
    tb.print_exception(e)
finally:
    book.close()
