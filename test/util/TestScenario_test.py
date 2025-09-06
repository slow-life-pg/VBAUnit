from pathlib import Path
import util.types as utypes


def test_scenario_singlegroup_singlemodule_basicproperty():
    s = utypes.TestScenario(Path("test\\util\\scenario_single_single.xlsx").resolve())
    assert s.valid
    assert s.count == 1
    assert s.path == Path("test\\util\\scenario_single_single.xlsx").resolve()


def test_scenario_singlegroup_singlemodule_groupinfo():
    s = utypes.TestScenario(Path("test\\util\\scenario_single_single.xlsx").resolve())
    assert s.count == 1
    g = s[0]
    assert g.groupname == "GroupA"
    assert s.group("GroupA") == g
    index = 0
    for g in s:
        index += 1
    assert index == 1


def test_scenario_singlegroup_singlemodule_moduleinf():
    s = utypes.TestScenario(Path("test\\util\\scenario_single_single.xlsx").resolve())
    assert s.count == 1
    g = s[0]
    assert g.count == 1
    m = g[0]
    assert m.testid == "A-001"
    assert m.subject == "Group Aのテストセット001"
    assert m.modulepath == Path("set001.py").resolve()
    assert m.run


def test_scenario_singlegroup_multimodule_basicproperty():
    s = utypes.TestScenario(Path("test\\util\\scenario_single_multi.xlsx").resolve())
    assert s.valid
    assert s.count == 1
    assert s.path == Path("test\\util\\scenario_single_multi.xlsx").resolve()
    assert s[0].count == 3


def test_scenario_singlegroup_multimodule_moduleinf():
    s = utypes.TestScenario(Path("test\\util\\scenario_single_multi.xlsx").resolve())
    assert s.count == 1
    g = s[0]
    assert g.count == 3
    m = g[0]
    assert m.testid == "A-001"
    assert m.subject == "Group Aのテストセット001"
    assert m.modulepath == Path("set001.py").resolve()
    assert m.run
    m = g[1]
    assert m.testid == "A-002"
    assert m.subject == "Group Aのテストセット002"
    assert m.modulepath == Path("set002.py").resolve()
    assert not m.run
    m = g[2]
    assert m.testid == "A-003"
    assert m.subject == "Group Aのテストセット003"
    assert m.modulepath == Path("set003.py").resolve()
    assert m.run


def test_scenario_multigroup_multimodule_basicproperty():
    s = utypes.TestScenario(Path("test\\util\\scenario_multi_multi.xlsx").resolve())
    assert s.valid
    assert s.count == 3
    assert s.path == Path("test\\util\\scenario_multi_multi.xlsx").resolve()
    assert s[0].count == 1
    assert s[0].groupname == "GroupA"
    assert s[1].count == 2
    assert s[1].groupname == "GroupB"
    assert s[2].count == 3
    assert s[2].groupname == "GroupC"
    gindex = 0
    for g in s:
        mindex = 0
        for m in g:
            mindex += 1
        gindex += 1
        assert mindex == gindex
    assert gindex == 3


def test_scenario_multigroup_multimodule_moduleinfo():
    s = utypes.TestScenario(Path("test\\util\\scenario_multi_multi.xlsx").resolve())
    assert s.valid
    assert s.count == 3
    g = s[0]
    assert g.count == 1
    m = g[0]
    assert m.testid == "A-001"
    assert m.subject == "Group Aのテストセット001"
    assert m.modulepath == Path("set001.py").resolve()
    assert m.run
    g = s[1]
    assert g.count == 2
    m = g[0]
    assert m.testid == "B-001"
    assert m.subject == "Group Bのテストセット001"
    assert m.modulepath == Path("set101.py").resolve()
    assert m.run
    m = g[1]
    assert m.testid == "B-002"
    assert m.subject == "Group Bのテストセット002"
    assert m.modulepath == Path("set102.py").resolve()
    assert m.run
    g = s[2]
    assert g.count == 3
    m = g[0]
    assert m.testid == "C-001"
    assert m.subject == "Group Cのテストセット001"
    assert m.modulepath == Path("set201.py").resolve()
    assert m.run
    m = g[1]
    assert m.testid == "C-002"
    assert m.subject == "Group Cのテストセット002"
    assert m.modulepath == Path("set202.py").resolve()
    assert m.run
    m = g[2]
    assert m.testid == "C-003"
    assert m.subject == "Group Cのテストセット003"
    assert m.modulepath == Path("set203.py").resolve()
    assert m.run
