from pathlib import Path
import util.types as utypes


def test_suite_singlegroup_singlemodule_nocoonstraints_basicproperty():
    sc = utypes.TestScenario(
        Path("test\\util\\scenario_single_single_loadable.xlsx").resolve()
    )
    s = utypes.TestSuite(
        "single_single_loadable",
        "test of single single noconstraints for loading testcases",
        sc,
    )
    assert s.name == "single_single_loadable"
    assert s.subject == "test of single single noconstraints for loading testcases"
    assert s.allscope
    assert not s.lastfailedscope
    assert s.count == 1
