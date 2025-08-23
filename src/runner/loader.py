from util.types import TestScenario, TestModule


class TestModuleLoader:
    """テストIDから該当するテストモジュールを動的に読み込む。"""

    def __init__(self, scenario: TestScenario):
        self.__scenario = scenario
        self.__last_executed: TestModule | None = None

    def load(self, testid: str) -> TestModule:
        """指定されたテストIDのモジュールをロードする"""
        testmodule = self.__get_module(testid)
        testmodule.load_module()
        self.__last_executed = testmodule
        return testmodule

    def unload(self) -> None:
        """最後に実行したテストモジュールを解放する。"""
        if self.__last_executed:
            self.__last_executed.unload_module()
            self.__last_executed = None

    def __get_module(self, testid: str) -> TestModule:
        for testgroup in self.__scenario:
            for module in testgroup:
                if module.testid == testid:
                    return module
        raise ValueError(f"TestModule with id {testid} not found")
