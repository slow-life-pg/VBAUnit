from pathlib import Path
from datetime import datetime
from collections import namedtuple
from dataclasses import dataclass, field
from enum import Enum
from pprint import pprint
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.descriptors.excel import CellRange


ResultCount = namedtuple("ResultCount", ["succeeded", "failed"])
GroupSet = namedtuple("GroupSet", ["spec", "result"])


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
    subject: str
    module: str

    @property
    def modulePath(self) -> Path:
        return Path(self.module).resolve()


@dataclass
class TestResult:
    testId: str
    succeeded: bool
    runnedAt: datetime


@dataclass
class TestGroup:
    groupName: str

    def __init__(self, name: str) -> None:
        self.groupName = name
        self.__testcases: list[TestCase] = []
        self.__results: dict[str, TestResult] = {}

    def addTestCase(self, testId: str, subject: str, module: str) -> None:
        testcase = TestCase(testId=testId, subject=subject, module=module)
        self.__testcases.append(testcase)

    def setResult(self, testId: str, succeeded: bool, runnedAt: datetime) -> None:
        succeeded = TestResult(testId=testId, succeeded=succeeded, runnedAt=runnedAt)
        self.__results[testId] = succeeded

    def getResult(self, testId: str) -> TestResult | None:
        if testId in self.__results:
            return self.__results[testId]
        else:
            return None

    def __iter__(self) -> list[TestCase]:
        return self.__testcases

    def __getitem__(self, index: int) -> TestCase | None:
        if index < 0 or len(self.__testcases) <= index:
            return None
        return self.__testcases[index]

    @property
    def results(self) -> list[TestResult]:
        return self.__results.values()

    @property
    def count(self) -> int:
        return len(self.__testcases)

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
        if not scenarioPath.exists():
            print(f"[ERR]file not exists. {str(scenarioPath)}")
            self.__valid = False
            return

        self.__valid = True
        self.__scenario = str(scenarioPath)
        self.__groups: list[TestGroup] = []
        self.__groupsIndex: dict[str, TestGroup] = {}
        self.__resultsIndex: dict[str, TestResult | None] = {}

        # ファイル読み込み
        sbook = None
        try:
            print(f"open scenario file: {self.__scenario}")
            sbook = load_workbook(self.__scenario, read_only=True)

            for gsheet in sbook.worksheets:
                group = self.__analyzegroup(gsheet=gsheet)
                if group is None:
                    print(f"[WARN]sheet '{gsheet.title}' is not valid test definition.")
                    continue
                self.__groups.append(group)
                self.__groupsIndex[group.groupName] = group
                self.__resultsIndex[group.groupName] = None

            if len(self.__groups) == 0:
                print(f"[ERR]No test found in '{self.__scenario}'.")
                self.__valid = False

        except Exception as ex:
            pprint(ex)
            self.__valid = False
            return
        finally:
            if sbook is not None:
                print("close scenario file.")
                sbook.close()

    def __iter__(self) -> list[TestGroup]:
        return self.__groups

    def __getitem__(self, index: int) -> TestGroup | None:
        if index < 0 or len(self.__groups) <= index:
            return None
        return self.__groups[index]

    @property
    def valid(self) -> bool:
        return self.__valid

    @property
    def count(self) -> int:
        return len(self.__groups)

    def group(self, groupName: str) -> tuple[TestCase, TestResult | None]:
        if groupName not in self.__groupsIndex:
            raise ValueError(f"[ERR]Unknown group name '{groupName}'.")

        return GroupSet(
            spec=self.__groupsIndex[groupName], result=self.__resultsIndex[groupName]
        )

    def __analyzegroup(self, gsheet: Worksheet) -> TestGroup | None:
        if gsheet.cell(2, 2).value != "テストID":
            return None
        if gsheet.cell(3, 2).value is None or gsheet.cell(3, 2).value == "":
            return None

        group = TestGroup(gsheet.title)
        row = 3
        while gsheet.cell(row, 2).value is not None and gsheet.cell(row, 2).value != "":
            group.addTestCase(
                testId=gsheet.cell(row, 2).value,
                subject=gsheet.cell(row, 3).value,
                module=gsheet.cell(row, 4).value,
            )
            row += 1

        return group
