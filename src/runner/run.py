from pathlib import Path
from datetime import datetime
from loader import TestModuleLoader
from util.types import Config, Scenario, TestSuite


def run_testsuite(config: Config, testsuitename: str) -> None:
    print(f"Running test suite for scenario: {config.scenario}")

    scenario = Scenario(scenariopath=Path(config.scenario))
    loader = TestModuleLoader(scenario=scenario)
    suite = __get_testsuite(config, testsuitename)


def __get_testsuite(config: Config, testsuitename: str) -> TestSuite:
    for suite in config:
        if suite.name == testsuitename:
            return suite
    raise ValueError(f"TestSuite with name {testsuitename} not found")
