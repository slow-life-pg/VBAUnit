from pathlib import Path
import shutil
import json
from datetime import datetime
from util.types import TestSuite, TestCase, TestResult
from openpyxl import load_workbook


def __gettimestampnow() -> str:
    return datetime.now().isoformat(sep=" ", timespec="milliesconds")


def __createresult(testcase: TestCase, succeeded: bool) -> TestResult:
    result = TestResult(
        testid=testcase.testid,
        group=testcase.group,
        module=str(testcase.module.modulepath),
        testfunction=testcase.testfunction,
        succeeded=succeeded,
        runned_at=__gettimestampnow(),
    )
    return result


def __runtestcase(testcase: TestCase) -> TestResult:
    return __createresult(testcase=testcase, succeeded=False)


def run_testsuite(suite: TestSuite, scenario: Path, bridge: Path, out: Path) -> None:
    print(f"Running test suite for scenario: {suite.name}")
    print(f"{suite.subject}")
    print(f"{scenario}")
    print(f"Output: {out}")

    outputpath = out.joinpath(scenario.name)
    shutil.copy(scenario, outputpath)

    # 実行

    results = list[TestResult]()
    for testcase in suite:
        config: dict[str, str] = {
            "testid": testcase.testid,
            "group": testcase.group,
            "module": testcase.module.modulepath.name,
            "function": testcase.testfunction,
            "start": __gettimestampnow(),
        }
        print("Running test case:")
        print(json.dumps(config))

        try:
            result = __runtestcase(testcase=testcase, bridge=bridge)
        except Exception as e:
            print(f"Runtime error: {e}")
            result = __createresult(testcase=testcase, succeeded=False)

        outcome: dict[str, str] = {
            "succeeded": str(result.succeeded),
            "end": result.runned_at,
        }
        print(json.dumps(outcome))
        results.append(result)

    # 結果の書込

    resultbook = None
    try:
        resultbook = load_workbook(outputpath)
    finally:
        if resultbook:
            resultbook.close()
