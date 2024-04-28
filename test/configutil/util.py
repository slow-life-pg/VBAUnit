import json
import shutil
from pathlib import Path
from enum import StrEnum
from pathutil.util import gettooldir


class ConfigPlace(StrEnum):
    LOCAL = "local"
    TOOL = "tool"
    OUTSIDE = "outside"


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
