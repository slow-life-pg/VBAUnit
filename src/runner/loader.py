from pathlib import Path
import openpyxl as xl


class ScenarioLoader:
    def __init__(self, scenariopath: Path) -> None:
        self.__scenario = scenariopath
