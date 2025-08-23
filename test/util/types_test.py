import pdb
import gc
from pathlib import Path
from util.types import Config, TestScope, TestScenario
from configutil.util import (
    ScenarioElement,
    createscenariofile,
    deletetestfiles,
)


def test_config_empty():
    configjson = {"scenario": "", "testsuites": []}
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 0


def test_config_senarioname():
    configjson = {"scenario": "scenario\\book.xlsx", "testsuites": []}
    config = Config(configjson)
    assert config.valid
    assert config.scenario == "scenario\\book.xlsx"
    assert len(config.testsuites) == 0


def test_config_seminormal_noscenario():
    configjson = {"testsuites": []}
    config = Config(configjson)
    assert not config.valid


def test_config_seminormal_notestsuites():
    configjson = {"scenario": ""}
    config = Config(configjson)
    assert not config.valid


def test_config_seminormal_noname():
    configjson = {"scenario": "", "testsuites": [{"tests": []}]}
    config = Config(configjson)
    assert not config.valid


def test_config_seminormal_noscopetests():
    configjson = {"scenario": "", "testsuites": [{"name": "testcase A"}]}
    config = Config(configjson)
    assert not config.valid


def test_config_seminormal_nogroup():
    configjson = {"scenario": "", "testsuites": [{"name": "testcase A", "tests": [{}]}]}
    config = Config(configjson)
    assert not config.valid


def test_config_singletestcase():
    configjson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "scope": "full"}],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "testcase A"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 0


def test_config_singletestcase2():
    configjson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "subject": "test for config with single testcase",
                "scope": "full",
            }
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "test for config with single testcase"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 0


def test_config_singletestcase3():
    configjson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "scope": "lastfailed"}],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "testcase A"
    assert config.testsuites[0].scope == TestScope.LASTFAILED
    assert len(config.testsuites[0].tests) == 0


def test_config_twotestcases():
    configjson = {
        "scenario": "",
        "testsuites": [
            {"name": "testcase A", "scope": "full"},
            {"name": "testcase B", "scope": "full"},
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 2
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "testcase A"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 0
    assert config.testsuites[1].name == "testcase B"
    assert config.testsuites[1].subject == "testcase B"
    assert config.testsuites[1].scope == TestScope.FULL
    assert len(config.testsuites[1].tests) == 0


def test_config_singletestset():
    configjson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "tests": [{"group": "Group A"}]}],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 0
    assert len(config.testsuites[0].tests[0].ignores) == 0


def test_config_singletestset2():
    configjson = {
        "scenario": "",
        "testsuites": [
            {"name": "testcase A", "tests": [{"group": "Group A", "filters": "warn"}]}
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 1
    assert config.testsuites[0].tests[0].filters[0] == "warn"
    assert len(config.testsuites[0].tests[0].ignores) == 0


def test_config_singletestset3():
    configjson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "tests": [{"group": "Group A", "filters": "warn | error"}],
            }
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 2
    assert config.testsuites[0].tests[0].filters[0] == "warn"
    assert config.testsuites[0].tests[0].filters[1] == "error"
    assert len(config.testsuites[0].tests[0].ignores) == 0


def test_config_singletestset4():
    configjson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "tests": [{"group": "Group A", "ignores": "warn | error"}],
            }
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 0
    assert len(config.testsuites[0].tests[0].ignores) == 2
    assert config.testsuites[0].tests[0].ignores[0] == "warn"
    assert config.testsuites[0].tests[0].ignores[1] == "error"


def test_config_twotestset():
    configjson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "tests": [
                    {"group": "Group A", "filters": "warn | error"},
                    {"group": "Group B", "ignores": "warn | error"},
                ],
            }
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 2
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 2
    assert config.testsuites[0].tests[0].filters[0] == "warn"
    assert config.testsuites[0].tests[0].filters[1] == "error"
    assert len(config.testsuites[0].tests[0].ignores) == 0
    assert config.testsuites[0].tests[1].group == "Group B"
    assert len(config.testsuites[0].tests[1].filters) == 0
    assert len(config.testsuites[0].tests[1].ignores) == 2
    assert config.testsuites[0].tests[1].ignores[0] == "warn"
    assert config.testsuites[0].tests[1].ignores[1] == "error"


def test_config_manytests():
    configjson = {
        "scenario": "c:\\dev\\excelvba\\target.xlsx",
        "testsuites": [
            {
                "name": "testcase A",
                "scope": "full",
                "tests": [
                    {"group": "Group A", "filters": "warn | error"},
                    {"group": "Group B", "ignores": "warn | error"},
                ],
            },
            {
                "name": "testcase B",
                "subject": "this is the testcase for function B",
                "tests": [
                    {"group": "Group A", "filters": "normal"},
                ],
            },
            {"name": "testcase failed", "scope": "lastfailed"},
        ],
    }
    config = Config(configjson)
    assert config.valid
    assert config.scenario == "c:\\dev\\excelvba\\target.xlsx"
    assert config.scenario == "c:\\dev\\excelvba\\target.xlsx"
    assert len(config.testsuites) == 3

    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 2
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 2
    assert config.testsuites[0].tests[0].filters[0] == "warn"
    assert config.testsuites[0].tests[0].filters[1] == "error"
    assert len(config.testsuites[0].tests[0].ignores) == 0
    assert config.testsuites[0].tests[1].group == "Group B"
    assert len(config.testsuites[0].tests[1].filters) == 0
    assert len(config.testsuites[0].tests[1].ignores) == 2
    assert config.testsuites[0].tests[1].ignores[0] == "warn"
    assert config.testsuites[0].tests[1].ignores[1] == "error"

    assert config.testsuites[1].name == "testcase B"
    assert config.testsuites[1].subject == "this is the testcase for function B"
    assert config.testsuites[1].scope == TestScope.FULL
    assert len(config.testsuites[1].tests) == 1
    assert config.testsuites[1].tests[0].group == "Group A"
    assert len(config.testsuites[1].tests[0].filters) == 1
    assert config.testsuites[1].tests[0].filters[0] == "normal"
    assert len(config.testsuites[1].tests[0].ignores) == 0

    assert config.testsuites[2].name == "testcase failed"
    assert config.testsuites[2].scope == TestScope.LASTFAILED
    assert len(config.testsuites[2].tests) == 0


def test_scenario_nofile():
    testPath = Path("scenario.xlsx").resolve()
    if testPath.exists():
        testPath.unlink()
    scenario = TestScenario(testPath)
    assert not scenario.valid


def test_scenario_nocontent():
    # デフォルトのシートが勝手に作られるので実質"notestcase"と同じ
    scenariopath = createscenariofile([])
    scenario = TestScenario(scenariopath=scenariopath)
    gc.collect()

    assert not scenario.valid

    deletetestfiles()


def test_scenario_notestcase():
    scenariopath = createscenariofile([("Group A", [])])
    scenario = TestScenario(scenariopath=scenariopath)
    gc.collect()

    assert not scenario.valid

    deletetestfiles()


def test_scenario_singletestcase():
    element1 = ScenarioElement(id="ID1", subject="Testcase 1", module="test1.py")
    scenariopath = createscenariofile([("Group A", [element1])])
    scenario = TestScenario(scenariopath=scenariopath)
    gc.collect()

    assert scenario.valid
    assert scenario.count == 1
    assert scenario[0].groupname == "Group A"
    assert scenario[0].count == 1
    assert scenario[0][0].testid == "ID1"
    assert scenario[0][0].subject == "Testcase 1"
    assert scenario[0][0].testid == "test1.py"

    deletetestfiles()


def test_scenario_multitestcase():
    element1 = ScenarioElement(id="ID1", subject="Testcase 1", module="test1.py")
    element2 = ScenarioElement(id="ID2", subject="Testcase 2", module="test2.py")
    element3 = ScenarioElement(id="ID3", subject="Testcase 3", module="test2.py")
    scenariopath = createscenariofile([("Group A", [element1, element2, element3])])
    scenario = TestScenario(scenariopath=scenariopath)
    gc.collect()

    assert scenario.valid
    assert scenario.count == 1
    assert scenario[0].groupname == "Group A"
    assert scenario[0].count == 3
    assert scenario[0][0].testid == "ID1"
    assert scenario[0][0].subject == "Testcase 1"
    assert scenario[0][0].testid == "test1.py"
    assert scenario[0][1].testid == "ID2"
    assert scenario[0][1].subject == "Testcase 2"
    assert scenario[0][1].testid == "test2.py"
    assert scenario[0][2].testid == "ID3"
    assert scenario[0][2].subject == "Testcase 3"
    assert scenario[0][2].testid == "test2.py"

    deletetestfiles()


def test_scenario_multigroup():
    element1 = ScenarioElement(id="ID1", subject="Testcase 1", module="test1.py")
    element2 = ScenarioElement(id="ID2", subject="Testcase 2", module="test2.py")
    element3 = ScenarioElement(id="ID3", subject="Testcase 3", module="test2.py")
    scenariopath = createscenariofile(
        [("Group A", [element1]), ("Group B", [element2]), ("Group C", [element3])]
    )
    scenario = TestScenario(scenariopath=scenariopath)
    gc.collect()

    assert scenario.valid
    assert scenario.count == 3
    assert scenario[0].groupname == "Group A"
    assert scenario[0].count == 1
    assert scenario[0][0].testid == "ID1"
    assert scenario[0][0].subject == "Testcase 1"
    assert scenario[0][0].testid == "test1.py"
    assert scenario[1].groupname == "Group B"
    assert scenario[1].count == 1
    assert scenario[1][0].testid == "ID2"
    assert scenario[1][0].subject == "Testcase 2"
    assert scenario[1][0].testid == "test2.py"
    assert scenario[2].groupname == "Group C"
    assert scenario[2].count == 1
    assert scenario[2][0].testid == "ID3"
    assert scenario[2][0].subject == "Testcase 3"
    assert scenario[2][0].testid == "test2.py"

    deletetestfiles()
