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
    testingTempDir = Path("testing").resolve()
    if not testingTempDir.exists():
        testingTempDir.mkdir()
    scenarioPath = testingTempDir.joinpath("scenario.xlsx")

    book = Workbook()
    defaultSheets = book.sheetnames

    tableId = 1
    for group in scenarioobj:
        print(f"create sheet: {group[0]}")
        sheet: Worksheet = book.create_sheet(group[0])
        sheet.cell(2, 2).value = "テストID"
        sheet.cell(2, 3).value = "説明"
        sheet.cell(2, 4).value = "モジュール"
        row = 3
        for testcase in group[1]:
            print(f"case: {testcase}")
            sheet.cell(row, 2).value = testcase.id
            sheet.cell(row, 3).value = testcase.subject
            sheet.cell(row, 4).value = testcase.module
            row += 1
        table = Table(id=tableId, ref=sheet.dimensions, displayName=f"Table{tableId}")
        sheet.tables.add(table=table)
        tableId += 1

    if len(book.worksheets) > len(defaultSheets):
        for title in defaultSheets:
            del book[title]
    print(f"sheets: {book.worksheets}")
    book.save(str(scenarioPath))
    book.close()

    return scenarioPath


<<<<<<< HEAD
def deletefile(filePath: Path) -> None:
    if filePath.exists():
        filePath.unlink()


=======
>>>>>>> 0f39a064ab17268f1757bd216b40f93a53ab7ac0
def deletetestfiles() -> None:
    testingTempDir = Path("testing").resolve()
    if testingTempDir.exists():
        shutil.rmtree(testingTempDir)


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
