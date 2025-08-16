from util.types import Scenario, TestModule


class TestModuleLoader:
    """テストIDから該当するテストモジュールを動的に読み込む。"""

    def __init__(self, scenario: Scenario):
        self.__scenario = scenario

    def load(self, testid: str) -> TestModule:
        """指定されたテストIDのモジュールをロードする"""
        testmodule = self.__get_module(testid)
        testmodule.load_module()
        return testmodule

    def unload(self) -> None:
        pass

    def __get_module(self, testid: str) -> TestModule:
        for testgroup in self.__scenario:
            for module in testgroup:
                if module.testid == testid:
                    return module
        raise ValueError(f"TestModule with id {testid} not found")
