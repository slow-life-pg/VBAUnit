from dataclasses import dataclass
import pytest
import util.types as utypes


@pytest.fixture
def patched_test_module(mocker):
    @dataclass
    class FakeTestModule:
        testid: str
        subject: str
        group: str
        modulepath: str
        run: bool

    mocker.patch("util.types.TestModule", FakeTestModule)
    # mocker.patch.object(utypes, "TestModule", FakeTestModule)

    return FakeTestModule


def test_init_and_count():
    g = utypes.TestGroup("GA")
    assert g.groupname == "GA"
    assert g.count == 0


def test_add_test_module(patched_test_module):
    g = utypes.TestGroup("GA")
    tm = g.add_test_module(
        testid="test1", subject="subject1", module="module1", run=True
    )
    assert g.count == 1
    assert g[0] == tm


def test_module_order(patched_test_module):
    g = utypes.TestGroup("GA")
    g.add_test_module(testid="test1", subject="subject1", module="module1", run=True)
    g.add_test_module(testid="test2", subject="subject2", module="module2", run=True)
    assert g.count == 2
    assert g[0].testid == "test1"
    assert g[1].testid == "test2"


def test_module_yield(patched_test_module):
    g = utypes.TestGroup("GA")
    tm = list[utypes.TestModule]()
    tm.append(
        g.add_test_module(
            testid="test1", subject="subject1", module="module1", run=True
        )
    )
    tm.append(
        g.add_test_module(
            testid="test2", subject="subject2", module="module2", run=True
        )
    )
    assert g.count == 2
    index = 0
    for m in g:
        assert m == tm[index]
        index += 1
