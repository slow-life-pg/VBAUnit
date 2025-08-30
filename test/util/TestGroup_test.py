import pytest
import pytest_mock as ptmocker
from dataclasses import dataclass
from pathlib import Path
import util.types as utypes


@pytest.fixture
def patched_TestModule(mocker):
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
