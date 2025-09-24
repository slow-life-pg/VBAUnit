from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_lookup01():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetLookupValue", "2", "", "")
        expect(result[0] is None)
        expect(result[2] == "鈴木")
        expect(result[3] == "花子")


def test_lookup02():
    testlib = gettestlib()
    with testlib.runapp("test/callmacro.xlsm"):
        result = testlib.callmacro(None, "GetLookupValue", "5", "", "")
        expect(result[0] is None)
        expect(result[2] == "山田")
        expect(result[3] == "晴人")
