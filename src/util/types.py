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
        except ValueError as ve:
            print("config parse value error:")
            pprint(ve)
            self.valid = False
        except Exception as e:
            print("config parse unknown error:")
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
        rawconditions = condition.split("|")
        for raw in rawconditions:
            conditions.append(raw.strip())


@dataclass
class TestCase:
    testid: str
    subject: str
    module: str

    @property
    def modulepath(self) -> Path:
        return Path(self.module).resolve()


@dataclass
class TestResult:
    testid: str
    succeeded: bool
    runned_at: datetime


@dataclass
class TestGroup:
    groupname: str

    def __init__(self, name: str) -> None:
        self.groupname = name
        self.__testcases: list[TestCase] = []
        self.__results: dict[str, TestResult] = {}

    def add_test_case(self, testdd: str, subject: str, module: str) -> None:
        testcase = TestCase(testid=testdd, subject=subject, module=module)
        self.__testcases.append(testcase)

    def set_result(self, testid: str, succeeded: bool, runned_at: datetime) -> None:
        succeeded = TestResult(testid=testid, succeeded=succeeded, runned_at=runned_at)
        self.__results[testid] = succeeded

    def get_result(self, testid: str) -> TestResult | None:
        if testid in self.__results:
            return self.__results[testid]
        else:
            return None

    def __iter__(self) -> list[TestCase]:
        return self

    def __next__(self) -> TestCase | None:
        if not self.__testcases:
            return None
        return self.__testcases.pop(0)

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
    def resultcount(self) -> ResultCount:
        success = 0
        fail = 0
        for result in self.__results.values():
            if result.succeeded:
                success += 1
            else:
                fail += 1
        return ResultCount(succeeded=success, failed=fail)


class Scenario:
    def __init__(self, scenariopath: Path) -> None:
        if not scenariopath.exists():
            print(f"[ERR]file not exists. {str(scenariopath)}")
            self.__valid = False
            return

        self.__valid = True
        self.__scenario = str(scenariopath)
        self.__groups: list[TestGroup] = []
        self.__groupsindex: dict[str, TestGroup] = {}
        self.__resultsindex: dict[str, TestResult | None] = {}

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
                self.__groupsindex[group.groupname] = group
                self.__resultsindex[group.groupname] = None

            if len(self.__groups) == 0:
                print(f"[ERR]No test found in '{self.__scenario}'.")
                self.__valid = False

        except ValueError as ve:
            print("Scenario parse value error :")
            pprint(ve)
            self.__valid = False
        except Exception as e:
            print("Scenario parse unknown error :")
            pprint(e)
            self.__valid = False
        finally:
            if sbook is not None:
                print("close scenario file.")
                sbook.close()

    def __iter__(self) -> list[TestGroup]:
        return self

    def __next__(self) -> TestGroup | None:
        if not self.__groups:
            return None
        return self.__groups.pop(0)

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

    def group(self, groupname: str) -> tuple[TestCase, TestResult | None]:
        if groupname not in self.__groupsindex:
            raise ValueError(f"[ERR]Unknown group name '{groupname}'.")

        return GroupSet(
            spec=self.__groupsindex[groupname], result=self.__resultsindex[groupname]
        )

    def __analyzegroup(self, gsheet: Worksheet) -> TestGroup | None:
        if gsheet.cell(2, 2).value != "テストID":
            return None
        if gsheet.cell(3, 2).value is None or gsheet.cell(3, 2).value == "":
            return None

        group = TestGroup(gsheet.title)
        row = 3
        while gsheet.cell(row, 2).value is not None and gsheet.cell(row, 2).value != "":
            group.add_test_case(
                testdd=gsheet.cell(row, 2).value,
                subject=gsheet.cell(row, 3).value,
                module=gsheet.cell(row, 4).value,
            )
            row += 1

        return group
