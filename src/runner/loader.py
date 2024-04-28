from pathlib import Path
import openpyxl as xl


class ScenarioLoader:
    def __init__(self, scenarioPath: Path) -> None:
        self.__scenario = scenarioPath
