from pathlib import Path
from datetime import datetime
from collections import namedtuple
from dataclasses import dataclass, field
from enum import Enum
from pprint import pprint
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


ResultCount = namedtuple("succeeded", "failed")


class TestScope(Enum):
    FULL = "full"
    LASTFAILED = "lastfailed"

    @classmethod
    def value_of(cls, target_value):
        for e in TestScope:
            if e.value == target_value:
                return e
        raise ValueError(f"TestScope: no member of value {target_value}")


@dataclass
class TestSet:
    group: str
    filters: list[str] = field(default_factory=list)
    ignores: list[str] = field(default_factory=list)


@dataclass
class TestSuite:
    name: str
    subject: str
    scope: TestScope = TestScope.FULL
    tests: list[TestSet] = field(default_factory=list)


class Config:
    def __init__(self, jsondict: dict) -> None:
        self.scenario = ""
        self.testsuites = list[TestSuite]()
        self.valid = True

        if "scenario" not in jsondict or "testsuites" not in jsondict:
            self.valid = False

        try:
            self.scenario = jsondict["scenario"]
            self.__parsetestsuites(testsuitesobj=jsondict["testsuites"])
        except Exception as e:
            print("config parse error:")
            pprint(e)
            self.valid = False

    def __parsetestsuites(self, testsuitesobj: list) -> None:
        for testsuiteobj in testsuitesobj:
            testsuite = self.__parsetestsuite(testsuiteobj=testsuiteobj)
            self.testsuites.append(testsuite)

    def __parsetestsuite(self, testsuiteobj: dict) -> TestSuite:
        if "name" not in testsuiteobj:
            raise ValueError("testsuite object must have 'name'", testsuiteobj)

        if "scope" not in testsuiteobj and "tests" not in testsuiteobj:
            raise ValueError(
                "testsuite object must have 'scope' or 'tests'", testsuiteobj
            )

        if "subject" in testsuiteobj:
            subject = testsuiteobj["subject"]
        else:
            subject = testsuiteobj["name"]

        if "scope" in testsuiteobj:
            scope = TestScope.value_of(testsuiteobj["scope"])
        else:
            scope = TestScope.FULL

        testsuite = TestSuite(name=testsuiteobj["name"], subject=subject, scope=scope)

        if "tests" in testsuiteobj:
            for testobj in testsuiteobj["tests"]:
                test = self.__parsetest(testobj=testobj)
                testsuite.tests.append(test)

        return testsuite

    def __parsetest(self, testobj: dict) -> TestSet:
        if "group" not in testobj:
            raise ValueError("test object must have 'group'", testobj)

        test = TestSet(group=testobj["group"])
        if "filters" in testobj:
            self.__parsecondition(testobj["filters"], test.filters)

        if "ignores" in testobj:
            self.__parsecondition(testobj["ignores"], test.ignores)

        return test

    def __parsecondition(self, condition: str, conditions: list[str]) -> None:
        rawConditions = condition.split("|")
        for raw in rawConditions:
            conditions.append(raw.strip())


@dataclass
class TestCase:
    testId: str
    module: str

    @property
    def modulePath(self) -> Path:
        return Path(self.module).resolve()


@dataclass
class TestResult:
    testId: str
    succeeded: bool
    runned: datetime


@dataclass
class TestGroup:
    groupName: str

    def __init__(self) -> None:
        self.__testcases: list[TestCase] = []
        self.__results: dict[str, TestResult] = {}

    def addTestCase(self, testId: str, module: str) -> None:
        testcase = TestCase(testId=testId, module=module)
        self.__testcases.append(testcase)

    def setResult(self, testId: str, succeeded: bool, runned: datetime) -> None:
        succeeded = TestResult(testId=testId, succeeded=succeeded, runned=runned)
        self.__results[testId] = succeeded

    def getResult(self, testId: str) -> TestResult | None:
        if testId in self.__results:
            return self.__results[testId]
        else:
            return None

    def __iter__(self) -> list[TestCase]:
        return self.__testcases

    @property
    def results(self) -> list[TestResult]:
        return self.__results.values()

    @property
    def resultCount(self) -> ResultCount:
        success = 0
        fail = 0
        for result in self.__results.values():
            if result.succeeded:
                success += 1
            else:
                fail += 1
        return ResultCount(succeeded=success, failed=fail)


class Scenario:
    def __init__(self, scenarioPath: Path) -> None:
        self.__valid = True
        self.__scenario = str(scenarioPath)
        self.__groups: list[TestGroup] = []

        # ファイル読み込み
        sbook = None
        try:
            sbook = Workbook(self.__scenario)

            for gsheet in sbook.worksheets:
                group = self.__analyzegroup(gsheet=gsheet)
                if group is None:
                    print(f"[WARN]sheet {gsheet.name} is not valid test definition.")
                    continue
                self.__groups.append(group)

        except Exception as ex:
            pprint(ex)
            return
        finally:
            if sbook is not None:
                sbook.close()

    def __iter__(self) -> list[TestGroup]:
        return self.__groups

    @property
    def valid(self) -> bool:
        return self.__valid

    def __analyzegroup(self, gsheet: Worksheet) -> TestGroup | None:
        pass
