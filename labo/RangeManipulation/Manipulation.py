import xlwings as xl

app = xl.App(visible=False)
book = app.books.open("Target.xlsm")

datasheet = book.sheets["DataSheet"]

# rangeに座標を与えるだけなら値が戻る
# rangeに範囲を与えるとリストが戻る
print(f"DataSheet A1: {datasheet.range('A1').value}")
print(f"DataSheet A1:E2: {datasheet.range('A1:E2').value}")
print()

# A1にすると項目名が入る
# 項目名が要らないなら次の行から始める
datatable_expand = datasheet.range("A2").options(expand="table").value
print(f"DataTable expand: {datatable_expand}")
print()

# テーブル名でrangeを指定すると先頭行を省いてデータだけ取れる
datatable_name = datasheet.range("SampleData").value
print(f"DataTable name: {datatable_name}")

book.close()
app.quit()
