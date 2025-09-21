# vbaunit_libをPathで通すパターン
import sys
from pathlib import Path

srcdir = Path(__file__).parent.parent.parent.joinpath("src")
sys.path.append(str(srcdir))

from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_rangevalue01():
    testlib = gettestlib()
    with testlib.runapp("test/rangevalue.xlsx") as wb:
        sheet1 = wb.sheets["Sheet1"]
        result = sheet1.range("A1:B1").value
        expect(result == ["A", "B"])
