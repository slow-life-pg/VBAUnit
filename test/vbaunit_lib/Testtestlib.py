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
    testbook = testlib.openexcel(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    )
    assert testbook is not None
    assert testbook.sheets.count == 1
    testlib.exitapp()


def test_get_simple_message():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    testbook = testlib.openexcel(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    )
    gsm = testbook.macro("GetSimpleMessage")
    msg = gsm()
    assert msg == "Simple Module Message"
    testlib.exitapp()


def test_get_simple_message_withmodulename():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    testbook = testlib.openexcel(
        "C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm"
    )
    gsm = testbook.macro("Module1.GetSimpleMessage")
    msg = gsm()
    assert msg == "Simple Module Message"
    testlib.exitapp()
