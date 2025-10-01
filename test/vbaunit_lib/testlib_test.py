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
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm") as testbook:
        assert testbook is not None
        assert testbook.sheets.count == 1


def test_get_simple_message():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm") as testbook:
        gsm = testbook.macro("GetSimpleMessage")
        msg = gsm()
        assert msg == "Simple Module Message"


def test_get_simple_message_withmodulename():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/GetSimpleMessage.xlsm") as testbook:
        gsm = testbook.macro("Module1.GetSimpleMessage")
        msg = gsm()
        assert msg == "Simple Module Message"


def test_passingaround_regex():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm") as testbook:
        rgx = testlib.getregexobj()
        gd = testbook.macro("GetDigits")
        msg = gd(rgx)
        testlib.freeobj(rgx)
        assert msg == "123456"


def test_passingaround_collection():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm") as testbook:
        coll = testlib.getcollectionobj()
        sl = testbook.macro("SetListValue")
        sl(coll, 2)
        assert coll.Count() == 2
        assert coll.item(1) == 2
        assert coll.item(2) == 4


def test_passingaround_dictionary():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/PassingObjectAround.xlsm") as testbook:
        expected = {2: "ABC", 4: "DEF"}
        dic = testlib.getdictionaryobj()
        sd = testbook.macro("SetDictionaryValue")
        sd(dic, 2)
        assert dic.Count == len(expected)
        for key in dic.Keys():
            assert dic.Item(key) == expected[key]


def test_callmacro_byref():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        p = 0  # dummy
        res = testlib.callmacro(None, "BackRefInt", p)
        assert len(res) == 4
        assert p == 0
        assert res[2] == 0
        assert res[0] is None
        assert res[1] == 100


def test_callmacro_return():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        p = 123
        res = testlib.callmacro(None, "BackReturn", p)
        assert len(res) == 4
        assert res[2] == 0
        assert res[0] == "Value is 123"
        assert res[1] == 123


def test_callmacro_createobject():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        res = testlib.callcreativemacro(None, "CreateClass1Object")
        print(f"Error: {res[1]} {res[2]}")
        assert len(res) == 3
        assert res[1] == 0
        comobj = res[0]
        assert comobj is not None


def test_object_property():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        res = testlib.callcreativemacro(None, "CreateClass2Object")
        print(f"Error: {res[1]} {res[2]}")
        assert len(res) == 3
        assert res[1] == 0
        comobj = res[0]
        assert comobj is not None
        assert comobj.Message == "This is Class2"


def test_object_function():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        res = testlib.callcreativemacro(None, "CreateClass2Object")
        print(f"Error: {res[1]} {res[2]}")
        assert len(res) == 3
        assert res[0] is not None
        msg = res[0].GetCustomMessage("CallMacro")
        assert msg == "This is CallMacro from Class2"


def test_object_byref():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/CallMacro.xlsm"):
        res = testlib.callcreativemacro(None, "CreateClass2Object")
        assert len(res) == 3
        assert res[0] is not None
        res2 = testlib.callmacro(res[0], "GetMessageLength", "CallMacro", 0)
        print(f"Error: {res2[3]} {res2[4]}")
        assert len(res2) == 5
        assert res2[3] == 0
        assert res2[0] is None
        assert res2[2] == 9


def test_dynamic_object_create():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/DynamicObject.xlsm"):
        obj = testlib.create_newinstance("Class1")
        assert obj is not None


def test_dynamic_object_directcall():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/DynamicObject.xlsm"):
        obj = testlib.create_newinstance("Class1")
        assert obj is not None

        obj.SetMessage("dynamic object")
        assert obj.GetMessage() == "dynamic object"


def test_dynamic_object_callmacro():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/DynamicObject.xlsm"):
        obj = testlib.create_newinstance("Class1")
        assert obj is not None

        testlib.callmacro(obj, "SetMessage", "dynamic object")
        res_class = testlib.callmacro(obj, "GetMessage")
        assert res_class[0] == "dynamic object"


def test_dynamic_object_pass():
    setglobalbridgepath(Path("C:/Dev/VBAUnit/src/VBAUnitCOMBridge.xlsm"))
    testlib = gettestlib()  # withapp=False, visible=False
    with testlib.runapp("C:/Dev/VBAUnit/test/vbaunit_lib/DynamicObject.xlsm"):
        obj = testlib.create_newinstance("Class1")
        assert obj is not None

        obj.SetMessage("dynamic object")
        res = testlib.callmacro(None, "GetObjectValue", obj)
        assert res[0] == "Message: dynamic object"

        res = testlib.callmacro(None, "GetValueLength", obj)
        assert res[0] == 14


def test_assertion_pass():
    from vbaunit_lib.testlib import expect

    try:
        expect(1 + 1 == 2)
    except AssertionError:
        assert False, "This should not happen"
    assert True


def test_assertion_fault():
    from vbaunit_lib.testlib import expect

    try:
        expect(1 + 1 == 3)
        assert False, "This should not happen"
    except AssertionError as ae:
        assert str(ae).startswith("Check failed at ")
        assert " -> expect(1 + 1 == 3)" in str(ae)
