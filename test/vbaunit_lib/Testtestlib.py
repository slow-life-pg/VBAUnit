from pathlib import Path
from vbaunit_lib.testlib import (
    setglobalbridgepath,
    getglobalbridgepath,
    gettestlib,
    VBAUnitTestLib,
)


def test_globalbridgepath():
    testpath = Path("C:/test/path/VBAUnitCOMBridge.xlsm")
    setglobalbridgepath(testpath)
    assert getglobalbridgepath() == testpath


def test_gettestlib():
    testlib = gettestlib()
    assert isinstance(testlib, VBAUnitTestLib)


def test_start_withoutapp():
    testlib = gettestlib(withapp=False)
    assert not testlib.appready


def test_start_withapp():
    testlib = gettestlib(withapp=True, visible=True)
    assert testlib.appready
    testlib.exitapp()


def test_open_close_book():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    ) as testbook:
        assert testbook is not None
        assert testbook.sheets.count == 1


def test_get_simple_message():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    ) as testbook:
        gsm = testbook.macro("GetSimpleMessage")
        msg = gsm()
        assert msg == "Simple Module Message"


def test_get_simple_message_withmodulename():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    ) as testbook:
        gsm = testbook.macro("Module1.GetSimpleMessage")
        msg = gsm()
        assert msg == "Simple Module Message"


def test_passingaround_regex():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm"
    ) as testbook:
        rgx = testlib.getregexobj()
        gd = testbook.macro("GetDigits")
        msg = gd(rgx)
        testlib.freeobj(rgx)
        assert msg == "123456"


def test_passingaround_collection():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm"
    ) as testbook:
        coll = testlib.getcollectionobj()
        sl = testbook.macro("SetListValue")
        sl(coll, 2)
        assert coll.Count() == 2
        assert coll.item(1) == 2
        assert coll.item(2) == 4


def test_passingaround_dictionary():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp(
        "C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm"
    ) as testbook:
        expected = {2: "ABC", 4: "DEF"}
        dic = testlib.getdictionaryobj()
        sd = testbook.macro("SetDictionaryValue")
        sd(dic, 2)
        assert dic.Count == len(expected)
        for key in dic.Keys():
            assert dic.Item(key) == expected[key]
