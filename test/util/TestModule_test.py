from pathlib import Path
from datetime import date
from util.types import ResultCount, TestModule, TestCase


def test_init_and_count():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    assert m.testid == "testidA"
    assert m.count == 0


def test_add_test_case():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    tc = m.add_testcase("func1", "func subject1", False)
    assert m.count == 1
    assert m[0] == tc


def test_case_order():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("func1", "func subject1", False)
    m.add_testcase("func2", "func subject2", False)
    assert m.count == 2
    assert m[0].testfunction == "func1"
    assert m[1].testfunction == "func2"


def test_module_yield():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    tc = list[TestCase]()
    tc.append(m.add_testcase("func1", "func subject1", False))
    tc.append(m.add_testcase("func2", "func subject2", False))
    assert m.count == 2
    index = 0
    for c in m:
        assert c == tc[index]
        index += 1


def test_result_none_initial_state():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    tr = m.set_result("f", True, None)
    assert tr is None


def test_result_none_doesnt_match_when_set():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f", "s", False)
    tr = m.set_result("ff", True, None)
    assert tr is None


def test_result_none_doesnt_match_when_get():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f", "s", False)
    tr = m.set_result("f", True, None)
    assert tr is not None
    tr = m.get_result("ff")
    assert tr is None


def test_result_single_case():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f", "s", False)
    ra = date(2025, 8, 31)
    tr = m.set_result("f", True, ra)
    assert m.resultcount == ResultCount(succeeded=1, failed=0)
    assert tr.testid == "testidA"
    assert tr.group == "groupA"
    assert tr.module == "moduleA"
    assert tr.testfunction == "f"
    assert tr.succeeded is True
    assert tr.runned_at == ra


def test_result_single_case_overwrite():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f", "s", False)
    raa = date(2025, 8, 31)
    rab = date(2025, 9, 1)
    m.set_result("f", True, raa)
    m.set_result("f", False, rab)
    assert m.resultcount == ResultCount(succeeded=0, failed=1)
    tr = m.get_result("f")
    assert tr.succeeded is False
    assert tr.runned_at == rab


def test_result_dual_case():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f1", "s", False)
    m.add_testcase("f2", "s", False)
    raa = date(2025, 8, 31)
    rab = date(2025, 9, 1)
    tr1 = m.set_result("f1", True, raa)
    tr2 = m.set_result("f2", False, rab)
    assert m.resultcount == ResultCount(succeeded=1, failed=1)
    assert tr1.testfunction == "f1"
    assert tr1.succeeded is True
    assert tr1.runned_at == raa
    assert tr2.testfunction == "f2"
    assert tr2.succeeded is False
    assert tr2.runned_at == rab


def test_result_dual_case_yield():
    m = TestModule("testidA", "subjectA", "groupA", "moduleA", True)
    m.add_testcase("f2", "s", False)
    m.add_testcase("f1", "s", False)
    raa = date(2025, 8, 31)
    rab = date(2025, 9, 1)
    tr2 = m.set_result("f2", False, rab)
    tr1 = m.set_result("f1", True, raa)
    assert m.resultcount == ResultCount(succeeded=1, failed=1)
    index = 0
    for r in m.results:
        if index == 0:
            assert r == tr1
        elif index == 1:
            assert r == tr2
        index += 1


def test_load_module_unloaded():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    assert m.modulepath == Path("test\\util\\loadee.py").resolve()
    assert m.testmodule is None


def test_load_module_loaded():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    m.load_module()
    assert m.testmodule is not None
    assert m.testmodule.return_test_str() == "This is test function in loadee module."
    m.unload_module()


def test_unload_module_unloaded():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    m.unload_module()
    assert m.testmodule is None


def test_unload_module_loaded():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    m.load_module()
    m.unload_module()
    assert m.testmodule is None


def test_load_module_twice():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    m.load_module()
    assert m.testmodule is not None
    m.unload_module()
    assert m.testmodule is None
    m.load_module()
    assert m.testmodule is not None
    assert m.testmodule.return_test_str() == "This is test function in loadee module."
    m.unload_module()
    assert m.testmodule is None


def test_pick_test_functions():
    m = TestModule("testidA", "subjectA", "groupA", "test\\util\\loadee.py", True)
    m.load_module()
    m.pick_testcases()
    m.unload_module()
    assert m.count == 2
    tgt: set[str] = set()
    for tc in m:
        tgt.add(tc.testfunction)
    assert tgt == {"test_function_run", "test_function_ignore"}
    print(f"run: {m['test_function_run'].ignore}")
    print(f"ignore: {m['test_function_ignore'].ignore}")
    assert not m["test_function_run"].ignore
    assert m["test_function_ignore"].ignore is True
