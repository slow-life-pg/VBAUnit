"""
既存のExcelに定義されたClassモジュールのインスタンスを得る実験

任意のClassを単体テストする方法を用意する。
"""

import os
import xlwings as xw

moduleDir = os.path.dirname(__file__)

app: xw.App = xw.App(visible=False)
book: xw.Book = app.books.open("classmodule.xlsm")
try:
    targetClassName = "cls_Class"
    tempComponent = book.api.VBProject.VBComponents.Add(1)  # 1:標準モジュール

    tempCode = (
        f"Public Function GetInstanceOf__{targetClassName}() As Variant\n"
        f"    Set GetInstanceOf__{targetClassName} = New {targetClassName}\n"
        " End Function\n"
    )
    tempComponent.CodeModule.AddFromString(tempCode)

    createModuleName = f"{tempComponent.Name}.GetInstanceOf__{targetClassName}"
    print(f"module name: {createModuleName}")
    createMacro = book.macro(createModuleName)
    instance = createMacro()

    book.api.VBProject.VBComponents.Remove(tempComponent)

    print(instance.GetMessage())

finally:
    book.close()
    app.kill()
