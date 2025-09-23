# vbaunit_libをPathで通すパターン
import sys
from pathlib import Path

srcdir = Path(__file__).parent.parent.parent.joinpath("src")
sys.path.append(str(srcdir))

from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_rangevalue02_writeread():
    testlib = gettestlib()
    with testlib.runapp("test/rangevalue.xlsx") as wb:
        sheet1 = wb.sheets["Sheet1"]
        buffer = sheet1["A1:B1"].value
        sheet2 = wb.sheets["Sheet2"]
        sheet2["A1"].value = buffer
        result = sheet2.range("A1:D1").value
        expect(result == ["A", "B", "C", "D"])
