import json
import shutil
from collections import namedtuple
from pathlib import Path
from enum import StrEnum
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from pathutil.util import gettooldir


class ConfigPlace(StrEnum):
    LOCAL = "local"
    TOOL = "tool"
    OUTSIDE = "outside"


ScenarioElement = namedtuple("ScenarioElement", ["id", "subject", "module"])


def getlocalscenariofilenam() -> str:
    return "localfile.xlsx"


def createconfigfile(place: ConfigPlace) -> Path:
    configPath = __getactualpath(place=place)
    jsonObj = {"scenario": getlocalscenariofilenam()}
    with open(str(configPath), mode="wt", encoding="utf-8") as fp:
        json.dump(obj=jsonObj, fp=fp)
    return configPath


def deleteconfigfile(place: ConfigPlace) -> None:
    configPath = __getactualpath(place=place)
    if configPath.exists():
        configPath.unlink()
    if place == ConfigPlace.OUTSIDE:
        configDir = configPath.parent
        shutil.rmtree(configDir, ignore_errors=True)


def createscenariofile(scenarioobj: list[tuple[str, list[ScenarioElement]]]) -> Path:
    scenarioPath = Path("scenario.xlsx").resolve()

    book = Workbook()

    tableId = 1
    for group in scenarioobj:
        sheet: Worksheet = book.create_sheet(group[0])
        rows = len(group[1])
        table = Table(id=tableId, ref=f"B2:D{rows+2}")
        sheet.tables.add(table=table)
        sheet.cell(2, 2).value = "テストID"
        sheet.cell(2, 3).value = "説明"
        sheet.cell(2, 4).value = "モジュール"
        row = 3
        for testcase in group[1]:
            sheet.cell(row, 2).value = testcase.id
            sheet.cell(row, 3).value = testcase.subject
            sheet.cell(row, 4).value = testcase.module
            row += 1
        tableId += 1

    book.save(str(scenarioPath))

    return scenarioPath


def deletescenariofile(scenarioPath: Path) -> None:
    if scenarioPath.exists():
        scenarioPath.unlink()


def __getactualpath(place: ConfigPlace) -> Path:
    if place == ConfigPlace.LOCAL:
        return __getlocalconfigpath()
    elif place == ConfigPlace.TOOL:
        return __gettoolconfigpath()
    elif place == ConfigPlace.OUTSIDE:
        return __getlocalconfigpath()
    else:
        return Path.cwd()


def __getlocalconfigpath() -> Path:
    localDir = Path.cwd()
    localConfigPath = localDir.joinpath("config.json")
    return localConfigPath


def __gettoolconfigpath() -> Path:
    toolDir = gettooldir()
    toolConfigPath = toolDir.joinpath("config.json")
    return toolConfigPath
