from pathlib import Path
from importlib import import_module
from datetime import datetime
from collections import namedtuple
from dataclasses import dataclass, field
from typing import Iterator
from enum import Enum
from pprint import pprint
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


ResultCount = namedtuple("ResultCount", ["succeeded", "failed"])
GroupSet = namedtuple("GroupSet", ["spec", "result"])


class TestScope(Enum):
    """どの範囲をテストするか
    FULL: グループに含まれるテストケース全て
    LASTFAILED: グループの中でも前回の実行で失敗したテストケースのみ
    """

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
    """あるグループの中でどのテストケースを実行するか定義するテストセット
    filters: 実行するべきテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り
    ignores: 無視するテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り
    """

    group: str
    filters: list[str] = field(default_factory=list)
    ignores: list[str] = field(default_factory=list)


@dataclass
class TestSuite:
    """まとめて実行するテストケースのまとまり
    name: テストスイートの名前。ログに書かれるのと、「前回の実行」を識別するもの
    subject: テストスイートの説明。任意
    scope: テストスイートの実行範囲。デフォルトはFULL
    tests: 実行するテストセットのリスト
    """

    name: str
    subject: str
    scope: TestScope = TestScope.FULL
    tests: list[TestSet] = field(default_factory=list)

    def append_test(self, group: str, filters: str = "", ignores: str = ""):
        """テストセットを追加する。
        group: テストグループ名
        filters: 実行するべきテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り。省略した場合はフィルタリングしない
        ignores: 無視するテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り。省略した場合は無視しない
        """
        test = TestSet(group=group, filters=[filters], ignores=[ignores])
        self.tests.append(test)


class Config:
    """テストの設定。config.jsonに定義される内容"""

    def __init__(self, jsondict: dict) -> None:
        self.scenario = ""
        self.testsuites = list[TestSuite]()
        self.valid = True

        self.__iteration = 0

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

    def __iter__(self) -> Iterator[TestSuite]:
        return self

    def __next__(self) -> TestSuite | None:
        if self.__iteration < len(self.testsuites):
            result = self.testsuites[self.__iteration]
            self.__iteration += 1
            return result

        self.__iteration = 0
        return None

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
    """個々のテストケース
    module: テストを実装したモジュールの相対パス
    subject: テストケースについての説明。任意
    testfunction: テストケースとして実行する関数名
    ignore: このテストケースを無視するかどうか
    """

    module: str
    subject: str
    testfunction: str
    ignore: bool


@dataclass
class TestResult:
    """テストケースの実行結果
    module: テストを実装したモジュールの相対パス
    testfunction: テストケースとして実行する関数名
    succeeded: テストケースが成功したかどうか
    runned_at: テストケースが実行された日時
    """

    module: str
    testfunction: str
    succeeded: bool
    runned_at: datetime


@dataclass
class TestModule:
    """テストを実装したモジュール
    testid: モジュール単位のテストID
    subject: モジュールについての説明。任意
    modulepath: モジュールの絶対パス
    """

    testid: str
    subject: str
    modulepath: str

    def __init__(self, testid: str, subject: str, modulepath: str) -> None:
        self.testid = testid
        self.subject = subject
        self.modulepath = modulepath

        self.__testcases: list[TestCase] = []
        self.__results: dict[str, TestResult] = {}

        self.__iterindex = 0

        self.__testmodule = None

    def load_module(self) -> None:
        self.__testmodule = import_module(self.modulepath)

    @property
    def testmodule(self):
        """ロードされたテストモジュールを返す"""
        return self.__testmodule

    def unload_module(self) -> None:
        del self.__testmodule

    def add_testcase(self, testfunction: str, subject: str, ignore: bool) -> None:
        self.__testcases.append(
            TestCase(
                module=self.testid,
                subject=subject,
                testfunction=testfunction,
                ignore=ignore,
            )
        )

    def set_result(
        self, testfunction: str, succeeded: bool, runned_at: datetime
    ) -> None:
        succeeded = TestResult(
            module=self.testid,
            testfunction=testfunction,
            succeeded=succeeded,
            runned_at=runned_at,
        )
        self.__results[self.__get_result_key(self.testid, testfunction)] = succeeded

    @property
    def modulepath(self) -> Path:
        return Path(self.testid).resolve()

    def get_result(self, testfunction: str) -> TestResult | None:
        if self.__get_result_key(self.testid, testfunction) in self.__results:
            return self.__results[self.__get_result_key(self.testid, testfunction)]
        else:
            return None

    def __get_result_key(self, module: str, testfunction: str) -> str:
        return testfunction + "@" + module

    @property
    def results(self) -> list[TestResult]:
        return self.__results.values()

    @property
    def resultcount(self) -> ResultCount:
        """成功と失敗の数を返す"""
        success = 0
        fail = 0
        for result in self.__results.values():
            if result.succeeded:
                success += 1
            else:
                fail += 1
        return ResultCount(succeeded=success, failed=fail)

    def __iter__(self) -> Iterator[TestCase]:
        """テストモジュールに含まれるテストケースのリスト"""
        return self

    def __next__(self) -> TestCase | None:
        if self.__iterindex >= len(self.__testcases):
            self.__iterindex = 0  # Reset for next iteration
            return None
        value = self.__testcases[self.__iterindex]
        self.__iterindex += 1
        return value

    def __getitem__(self, index: int | str) -> TestCase | None:
        if isinstance(index, int):
            if index < 0 or len(self.__testcases) <= index:
                return None
            return self.__testcases[index]
        else:
            for testcase in self.__testcases:
                if testcase.module == index:
                    return testcase
            return None

    @property
    def count(self) -> int:
        return len(self.__testcases)


@dataclass
class TestGroup:
    """意味的にまとまったテストケースのグループ。シナリオファイルの1つのシートに対応する
    groupname: グループの名前
    """

    groupname: str

    def __init__(self, name: str) -> None:
        """グループの初期化
        name: グループ名。シナリオファイルのシート名
        """
        self.groupname = name
        self.__testmodules: list[TestModule] = []

        self.__iterindex = 0

    def add_test_module(self, testid: str, module: str, subject: str) -> None:
        testmodule = TestModule(testid=testid, subject=subject, modulepath=module)
        self.__testmodules.append(testmodule)

    def __iter__(self) -> Iterator[TestModule]:
        """グループにかかわらず全ての定義されたテストモジュールのリスト"""
        return self

    def __next__(self) -> TestModule | None:
        if self.__iterindex >= len(self.__testmodules):
            self.__iterindex = 0  # Reset for next iteration
            return None
        value = self.__testmodules[self.__iterindex]
        self.__iterindex += 1
        return value

    def __getitem__(self, index: int | str) -> TestModule | None:
        if isinstance(index, int):
            if index < 0 or len(self.__testmodules) <= index:
                return None
            return self.__testmodules[index]
        else:
            for module in self.__testmodules:
                if module.testid == index:
                    return module
            return None

    @property
    def count(self) -> int:
        return len(self.__testmodules)


class Scenario:
    """テストシナリオの定義。シナリオファイルに対応する"""

    def __init__(self, scenariopath: Path) -> None:
        """テストシナリオの初期化
        scenariopath: シナリオファイルの絶対パス
        """
        if not scenariopath.exists():
            print(f"[ERR]file not exists. {str(scenariopath)}")
            self.__valid = False
            return

        self.__valid = True
        self.__scenario = str(scenariopath)
        self.__groups: list[TestGroup] = []
        self.__groupsindex: dict[str, TestGroup] = {}
        self.__resultsindex: dict[str, TestResult | None] = {}

        self.__iterindex = 0

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

    def __iter__(self) -> Iterator[TestGroup]:
        return self

    def __next__(self) -> TestGroup | None:
        if self.__iterindex >= len(self.__groups):
            self.__iterindex = 0  # Reset for next iteration
            return None
        value = self.__groups[self.__iterindex]
        self.__iterindex += 1
        return value

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

    @property
    def path(self) -> str:
        return self.__scenario

    def group(self, groupname: str) -> tuple[TestModule, TestResult | None]:
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
            group.add_test_module(
                testid=gsheet.cell(row, 2).value,
                subject=gsheet.cell(row, 3).value,
                module=gsheet.cell(row, 4).value,
            )
            row += 1

        return group
