import pdb
import gc
from pathlib import Path
from util.types import Config, TestScope, Scenario
from configutil.util import (
    ScenarioElement,
    createscenariofile,
    deletetestfiles,
)


def test_Config_empty():
    configJson = {"scenario": "", "testsuites": []}
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 0


def test_Config_senarioname():
    configJson = {"scenario": "scenario\\book.xlsx", "testsuites": []}
    config = Config(configJson)
    assert config.valid
    assert config.scenario == "scenario\\book.xlsx"
    assert len(config.testsuites) == 0


def test_Config_seminormal_noscenario():
    configJson = {"testsuites": []}
    config = Config(configJson)
    assert not config.valid


def test_Config_seminormal_notestsuites():
    configJson = {"scenario": ""}
    config = Config(configJson)
    assert not config.valid


def test_Config_seminormal_noname():
    configJson = {"scenario": "", "testsuites": [{"tests": []}]}
    config = Config(configJson)
    assert not config.valid


def test_Config_seminormal_noscopetests():
    configJson = {"scenario": "", "testsuites": [{"name": "testcase A"}]}
    config = Config(configJson)
    assert not config.valid


def test_Config_seminormal_nogroup():
    configJson = {"scenario": "", "testsuites": [{"name": "testcase A", "tests": [{}]}]}
    config = Config(configJson)
    assert not config.valid


def test_Config_singletestcase():
    configJson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "scope": "full"}],
    }
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "testcase A"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 0


def test_Config_singletestcase2():
    configJson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "subject": "test for config with single testcase",
                "scope": "full",
            }
        ],
    }
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "test for config with single testcase"
    assert config.testsuites[0].scope == TestScope.FULL
    assert len(config.testsuites[0].tests) == 0


def test_Config_singletestcase3():
    configJson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "scope": "lastfailed"}],
    }
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert config.testsuites[0].subject == "testcase A"
    assert config.testsuites[0].scope == TestScope.LASTFAILED
    assert len(config.testsuites[0].tests) == 0


def test_Config_twotestcases():
    configJson = {
        "scenario": "",
        "testsuites": [
            {"name": "testcase A", "scope": "full"},
            {"name": "testcase B", "scope": "full"},
        ],
    }
    config = Config(configJson)
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


def test_Config_singletestset():
    configJson = {
        "scenario": "",
        "testsuites": [{"name": "testcase A", "tests": [{"group": "Group A"}]}],
    }
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 0
    assert len(config.testsuites[0].tests[0].ignores) == 0


def test_Config_singletestset2():
    configJson = {
        "scenario": "",
        "testsuites": [
            {"name": "testcase A", "tests": [{"group": "Group A", "filters": "warn"}]}
        ],
    }
    config = Config(configJson)
    assert config.valid
    assert config.scenario == ""
    assert len(config.testsuites) == 1
    assert config.testsuites[0].name == "testcase A"
    assert len(config.testsuites[0].tests) == 1
    assert config.testsuites[0].tests[0].group == "Group A"
    assert len(config.testsuites[0].tests[0].filters) == 1
    assert config.testsuites[0].tests[0].filters[0] == "warn"
    assert len(config.testsuites[0].tests[0].ignores) == 0


def test_Config_singletestset3():
    configJson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "tests": [{"group": "Group A", "filters": "warn | error"}],
            }
        ],
    }
    config = Config(configJson)
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


def test_Config_singletestset4():
    configJson = {
        "scenario": "",
        "testsuites": [
            {
                "name": "testcase A",
                "tests": [{"group": "Group A", "ignores": "warn | error"}],
            }
        ],
    }
    config = Config(configJson)
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


def test_Config_twotestset():
    configJson = {
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
    config = Config(configJson)
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


def test_Config_manytests():
    configJson = {
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
    config = Config(configJson)
    assert config.valid
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


def test_Scenario_nofile():
    testPath = Path("scenario.xlsx").resolve()
    if testPath.exists():
        testPath.unlink()
    scenario = Scenario(testPath)
    assert not scenario.valid


def test_Scenario_nocontent():
    # デフォルトのシートが勝手に作られるので実質"notestcase"と同じ
    scenarioPath = createscenariofile([])
    scenario = Scenario(scenarioPath=scenarioPath)
    gc.collect()

    assert not scenario.valid

    deletetestfiles()


def test_Scenario_notestcase():
    scenarioPath = createscenariofile([("Group A", [])])
    scenario = Scenario(scenarioPath=scenarioPath)
    gc.collect()

    assert not scenario.valid

    deletetestfiles()


def test_Scenario_singletestcase():
    element1 = ScenarioElement(id="ID1", subject="Testcase 1", module="test1.py")
    scenarioPath = createscenariofile([("Group A", [element1])])
    scenario = Scenario(scenarioPath=scenarioPath)
    gc.collect()

    assert scenario.valid
    assert scenario.count == 1
    assert scenario[0].groupName == "Group A"
    assert scenario[0].count == 1

    deletetestfiles()
