from pathlib import Path
from datetime import datetime
from loader import TestModuleLoader
from util.types import TestScenario, TestSuite


def run_testsuite(suite: TestSuite, out: Path) -> None:
    print(f"Running test suite for scenario: {suite.name}")
