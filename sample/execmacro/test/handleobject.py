from vbaunit_lib.testlib import gettestlib, expect, expect_collection, expect_collection_list, expect_dictionary


def test_handling_collection():
    testlib = gettestlib()
    with testlib.runapp("test/handleobject.xlsm"):
        coll = testlib.getcollectionobj()
        result = testlib.callmacro(None, "FillCollection", coll)
        expect(result[0] is None)
        expect(coll.Count() == 3)
        expect_collection(coll, "1", "太郎")
        expect_collection(coll, "2", "次郎")

        expectation = ["太郎", "次郎", "花子"]
        expect_collection_list(coll, expectation)


def test_handling_dictionary():
    testlib = gettestlib()
    with testlib.runapp("test/handleobject.xlsm"):
        dic = testlib.getdictionaryobj()
        result = testlib.callmacro(None, "FillDictionary", dic)
        expect(result[0] is None)
        expect(dic.Count == 3)

        expectation = {1: "太郎", 2: "次郎", 3: "花子"}
        expect_dictionary(dic, expectation)
