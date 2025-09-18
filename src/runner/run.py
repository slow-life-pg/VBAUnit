from pathlib import Path
import shutil
import json
from datetime import datetime
from openpyxl import load_workbook
from util.types import TestSuite, TestCase, TestResult


def __gettimestampnow() -> str:
    return datetime.now().isoformat(sep=" ", timespec="milliesconds")


def __writestartlog(testcase: TestCase, f) -> None:
    config: dict[str, str] = {
        "testid": testcase.testid,
        "group": testcase.group,
        "module": testcase.module.modulepath.name,
        "function": testcase.testfunction,
        "at": str(testcase.start_line),
        "start": __gettimestampnow(),
    }
    f.write(json.dumps(config) + "\n")
    print("Running test case:")
    print(json.dumps(config))


def __writeendlog(result: TestResult, f) -> None:
    outcome: dict[str, str] = {
        "outcome": str(result.succeeded),
        "runned": result.runned_at,
    }
    f.write(json.dumps(outcome) + "\n")
    print(json.dumps(outcome))


def __createresult(testcase: TestCase, succeeded: bool) -> TestResult:
    result = TestResult(
        testid=testcase.testid,
        group=testcase.group,
        module=testcase.module,
        testfunction=testcase.testfunction,
        start_line=testcase.start_line,
        succeeded=succeeded,
        runned_at=__gettimestampnow(),
    )
    return result


def __runtestcase(testcase: TestCase) -> TestResult:
    return __createresult(testcase=testcase, succeeded=False)


def run_testsuite(suite: TestSuite, scenario: Path, bridge: Path, out: Path) -> None:
    outputpath = out.joinpath(scenario.name)
    testlogpath = out.joinpath("testlog.txt")

    with open(testlogpath, mode="w", encoding="utf-8") as f:
        f.write(f"Test log for {suite.name}\n")
        f.write(f"Subject: {suite.subject}\n")
        f.write(f"Scenario: {scenario}\n")
        f.write(f"Output: {out}\n")
        f.write(f"Bridge: {bridge}\n")
        f.write(f"Started at: {__gettimestampnow()}\n")
        f.write("\n")
    print(f"Running test suite for scenario: {suite.name}")
    print(f"{suite.subject}")
    print(f"{scenario}")
    print(f"Output: {out}")

    shutil.copy(scenario, outputpath)

    # 実行

    results = list[TestResult]()
    modulesummary_success: dict[str, int] = {}
    modulesummary_failure: dict[str, int] = {}
    with open(testlogpath, mode="a", encoding="utf-8") as f:
        for testcase in suite:
            __writestartlog(testcase=testcase, f=f)

            try:
                result = __runtestcase(testcase=testcase)
            except Exception as e:
                print(f"Runtime error: {e}")
                result = __createresult(testcase=testcase, succeeded=False)

            __writeendlog(result=result, f=f)
            f.flush()

            results.append(result)

            modulename = testcase.module.modulepath.name
            if modulename not in modulesummary_success:
                modulesummary_success[modulename] = 0
                modulesummary_failure[modulename] = 0
            if result.succeeded:
                modulesummary_success[modulename] += 1
            else:
                modulesummary_failure[modulename] += 1

    # 結果の書込

    resultbook = None
    try:
        resultbook = load_workbook(outputpath)

        # モジュールの結果

        # テストケースの結果

        resultbook.save(outputpath)
    finally:
        if resultbook:
            resultbook.close()
