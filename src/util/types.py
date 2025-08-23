from pathlib import Path
from importlib import import_module
from datetime import datetime
from collections import namedtuple
from dataclasses import dataclass, field
from typing import Iterator
from pprint import pprint
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


ResultCount = namedtuple("ResultCount", ["succeeded", "failed"])
GroupSet = namedtuple("GroupSet", ["spec", "result"])


@dataclass
class TestCase:
    """個々のテストケース
    testid: テストモジュールの識別ID
    module: テストを実装したモジュールの相対パス
    subject: テストケースについての説明。任意
    testfunction: テストケースとして実行する関数名
    ignore: このテストケースを無視するかどうか
    """

    testid: str
    module: str
    subject: str
    testfunction: str
    ignore: bool


@dataclass
class TestResult:
    """テストケースの実行結果
    testid: テストモジュールの識別ID
    module: テストを実装したモジュールの相対パス
    testfunction: テストケースとして実行する関数名
    succeeded: テストケースが成功したかどうか
    runned_at: テストケースが実行された日時
    """

    testid: str
    module: str
    testfunction: str
    succeeded: bool
    runned_at: datetime


@dataclass
class TestModule:
    """テストを実装したモジュール"""

    def __init__(self, testid: str, subject: str, modulepath: str, run: bool) -> None:
        """テストモジュールを作成する。
        testid: テストモジュールの識別ID
        subject: テストモジュールの説明
        modulepath: テストモジュールのパス
        run: テストモジュールを実行するかどうか
        """
        self.__testid = testid
        self.__subject = subject
        self.__modulepath = modulepath
        self.__run = run

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

    @property
    def testid(self):
        return self.__testid

    @property
    def subject(self):
        return self.__subject

    @property
    def modulepath(self) -> Path:
        return Path(self.__modulepath).resolve()

    @property
    def run(self) -> bool:
        return self.__run

    def unload_module(self) -> None:
        del self.__testmodule

    def add_testcase(self, testfunction: str, subject: str, ignore: bool) -> None:
        self.__testcases.append(
            TestCase(
                testid=self.__testid,
                module=self.__modulepath,
                subject=subject,
                testfunction=testfunction,
                ignore=ignore,
            )
        )

    def set_result(
        self, testfunction: str, succeeded: bool, runned_at: datetime
    ) -> None:
        succeeded = TestResult(
            testid=self.__testid,
            module=self.__modulepath,
            testfunction=testfunction,
            succeeded=succeeded,
            runned_at=runned_at,
        )
        self.__results[self.__get_result_key(self.__testid, testfunction)] = succeeded

    def get_result(self, testfunction: str) -> TestResult | None:
        if self.__get_result_key(self.__testid, testfunction) in self.__results:
            return self.__results[self.__get_result_key(self.__testid, testfunction)]
        else:
            return None

    def __get_result_key(self, testid: str, testfunction: str) -> str:
        return testfunction + "@" + testid

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
    """意味的にまとまったテストケースのグループ。シナリオファイルの1つのシートに対応する"""

    def __init__(self, name: str) -> None:
        """グループの初期化
        name: グループ名。シナリオファイルのシート名
        """
        self.__groupname = name
        self.__testmodules: list[TestModule] = []

        self.__iterindex = 0

    @property
    def groupname(self) -> str:
        return self.__groupname

    def add_test_module(
        self, testid: str, module: str, subject: str, run: bool
    ) -> None:
        testmodule = TestModule(
            testid=testid, subject=subject, modulepath=module, run=run
        )
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


class TestScenario:
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
        if gsheet.cell(3, 2).value is None or gsheet.cell(3, 2).value != "説明":
            return None
        if gsheet.cell(4, 2).value is None or gsheet.cell(4, 2).value != "モジュール":
            return None
        if gsheet.cell(5, 2).value is None or gsheet.cell(5, 2).value != "実行":
            return None
        if gsheet.cell(6, 2).value is None or gsheet.cell(6, 2).value != "結果":
            return None

        group = TestGroup(gsheet.title)
        row = 3
        while gsheet.cell(row, 2).value is not None and gsheet.cell(row, 2).value != "":
            group.add_test_module(
                testid=gsheet.cell(row, 2).value,
                subject=gsheet.cell(row, 3).value,
                module=gsheet.cell(row, 4).value,
                run=self.__getrequired(gsheet.cell(row, 5).value),
            )
            row += 1

        return group

    def __getrequired(self, runvalue: str) -> bool:
        return runvalue == "○"


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
    tests: list[TestCase] = field(default_factory=list)

    def __init__(
        self,
        name: str,
        subject: str,
        scenario: TestScenario,
        filters: str = "",
        ignores: str = "",
        scope: str = "all",
    ):
        """テストセットを作成する。
        name: テストスイートを識別する名称（この名前でログを出力する）
        subject: テストスイートの説明
        scenario: テストシナリオ
        filters: 実行するべきテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り。省略した場合はフィルタリングしない
        ignores: 無視するテストケースの関数名に含まれる識別子（任意の文字列）。"|"区切り。省略した場合は無視しない
        scope: 実行する範囲。全部か前回の失敗のみ
        """
        self.__name = name
        self.__subject = subject
        self.__filters = [f.strip() for f in filters.split("|")]
        self.__ignores = [i.strip() for i in ignores.split("|")]
        self.__scope = scope
        for group in scenario:
            for module in group:
                if not module.run:
                    continue
                for testcase in module:
                    if self.__torun(
                        testcase=testcase,
                        filters=self.__filters,
                        ignores=self.__ignores,
                    ):
                        self.tests.append(testcase)

    @property
    def name(self) -> str:
        return self.__name

    @property
    def subject(self) -> str:
        return self.__subject

    @property
    def allscope(self) -> bool:
        return self.__scope == "all"

    @property
    def lastfailedscope(self) -> bool:
        return self.__scope == "lastfailed"

    def __torun(
        self, testcase: TestCase, filters: list[str], ignores: list[str]
    ) -> bool:
        if testcase.ignore:
            return False
        if self.__includestest(testcase.testfunction, ignores):
            return False
        if len(filters) == 0 or self.__includestest(testcase.testfunction, filters):
            return True

        return False

    def __includestest(self, testname: str, condition: list[str]) -> bool:
        if not condition or len(condition) == 0:
            return False
        for f in condition:
            if f in testname:
                return True
        return False
