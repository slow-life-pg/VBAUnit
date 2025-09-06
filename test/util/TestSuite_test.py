from pathlib import Path
import util.types as utypes


def test_suite_singlegroup_singlemodule_nocoonstraints_basicproperty():
    sc = utypes.TestScenario(Path("test\\util\\scenario_single_single.xlsx").resolve())
    s = utypes.TestSuite(
        "single_single", "test of single single noconstraints basic", sc
    )
    assert s.name == "single_single"
    assert s.subject == "test of single single noconstraints basic"
    assert s.allscope
    assert not s.lastfailedscope
    assert s.count == 1
