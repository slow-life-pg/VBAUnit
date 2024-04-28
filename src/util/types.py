from pathlib import Path
from datetime import datetime
from collections import namedtuple
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


class TestSet:
    def __init__(self) -> None:
        self.group = ""
        self.filters = list[str]()
        self.ignores = list[str]()


class TestSuite:
    def __init__(self) -> None:
        self.name = ""
        self.subject = ""
        self.scope = TestScope.FULL
        self.tests = list[TestSet]()


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

        testsuite = TestSuite()
        testsuite.name = testsuiteobj["name"]

        if "subject" in testsuiteobj:
            testsuite.subject = testsuiteobj["subject"]
        else:
            testsuite.subject = testsuite.name

        if "scope" in testsuiteobj:
            testsuite.scope = TestScope.value_of(testsuiteobj["scope"])

        if "tests" in testsuiteobj:
            for testobj in testsuiteobj["tests"]:
                test = self.__parsetest(testobj=testobj)
                testsuite.tests.append(test)

        return testsuite

    def __parsetest(self, testobj: dict) -> TestSet:
        if "group" not in testobj:
            raise ValueError("test object must have 'group'", testobj)

        test = TestSet()
        test.group = testobj["group"]
        if "filters" in testobj:
            self.__parsecondition(testobj["filters"], test.filters)

        if "ignores" in testobj:
            self.__parsecondition(testobj["ignores"], test.ignores)

        return test

    def __parsecondition(self, condition: str, conditions: list[str]) -> None:
        rawConditions = condition.split("|")
        for raw in rawConditions:
            conditions.append(raw.strip())


class TestCase:
    def __init__(self, testId: str, module: str) -> None:
        self.__id = testId
        self.__module: str = module
        self.__modulePath = Path(module).resolve()

    @property
    def testid(self) -> str:
        return self.__id

    @property
    def module(self) -> Path:
        return self.__modulePath


class TestResult:
    def __init__(self, testId: str, succeeded: bool, runned: datetime) -> None:
        self.__id = testId
        self.__succeeded = succeeded
        self.__runned = runned

    @property
    def testid(self) -> str:
        return self.__id

    @property
    def succeeded(self) -> bool:
        return self.__succeeded

    @property
    def runned(self) -> datetime:
        return self.__runned


class TestGroup:
    def __init__(self, name: str) -> None:
        self.__groupName = name
        self.__testcases: list[TestCase] = []
        self.__results: dict[str, TestResult] = {}

    @property
    def groupName(self) -> str:
        return self.__groupName

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
