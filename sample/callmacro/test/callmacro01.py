# vbaunit_libのパスは通っている想定のパターン
from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_getcellvalue01_ret():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetCellValueRet", 2, 3)  # 名前付きにすると最後の可変長引数で怒られる
        expect(result[0] == "C2")


def test_getcellvalue02_ret():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetCellValueRet", 7, 9)
        expect(result[0] == "I7")


def test_getcellvalue03_ref():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetCellValueRef", 2, 3, "")
        expect(result[0] is None)
        expect(result[3] == "C2")


def test_getcellvalue04_ref():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetCellValueRef", 9, 4, "")
        expect(result[0] is None)
        expect(result[3] == "D9")
