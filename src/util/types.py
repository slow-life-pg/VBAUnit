from enum import Enum
from pprint import pprint


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
    group: str
    filters: list[str] = []
    ignores: list[str] = []


class TestSuite:
    name: str
    subject: str
    scope: TestScope = TestScope.FULL
    tests: list[TestSet] = []


class Config:
    scenario: str
    testsuites: list[TestSuite] = []

    valid = True

    def __init__(self, jsondict: dict) -> None:
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
        return conditions
