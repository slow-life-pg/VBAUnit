from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_update01():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm") as wb:
        testlib.callmacro(None, "UpdateSheet")
        result = wb.sheets("Sheet2")["A1:D4"].value
        expect(result == [[1, 1, 1, None], [1, 1, 1, None], [1, 1, 1, None], [None, None, None, None]])
