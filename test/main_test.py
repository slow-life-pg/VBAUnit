import json
from pathlib import Path
import main

# このやり方もあるにはある　→　今回はconftest.pyで行ってみる
# import pytest
# @pytest.fixture
# def main_fixture():
#     # ここで検索パスを追加して
#     srcDir = Path(__file__).parent.parent.joinpath("src")
#     sys.path.append(str(srcDir))

#     # ここでimportしてもテストケースの実行時にModuleNotFoundError

#     yield

#     sys.path.remove(str(srcDir))


def gettooldir() -> Path:
    return Path(__file__).parent.parent.joinpath("src")  # コピペ注意！


def getlocalconfigpath() -> Path:
    localDir = Path.cwd()
    localConfigPath = localDir.joinpath("config.json")
    return localConfigPath


def gettoolconfigpath() -> Path:
    toolDir = gettooldir()
    toolConfigPath = toolDir.joinpath("config.json")
    return toolConfigPath


def createconfigfile(configPath: Path) -> None:
    jsonObj = {"scenario": "localfile.xlsx"}
    with open(str(configPath), "wt", encoding="utf-8") as fp:
        json.dump(obj=jsonObj, fp=fp)


def deleteconfigfile(configPath: Path) -> None:
    if configPath.exists():
        configPath.unlink()


def createlocalconfg():
    localConfigPath = getlocalconfigpath()
    createconfigfile(localConfigPath)


def deletelocalconfig():
    localConfigPath = getlocalconfigpath()
    deleteconfigfile(localConfigPath)


def createtoolconfig():
    toolConfigPath = gettoolconfigpath()
    createconfigfile(toolConfigPath)


def deletetoolconfig():
    toolConfigPath = gettoolconfigpath()
    deleteconfigfile(toolConfigPath)


def test_getconfigpath_incurrent():
    curDir = Path.cwd()
    toolDir = gettooldir()
    createlocalconfg()
    try:
        expected = curDir.joinpath("config.json")
        configPath = main.getconfigpath(currentDir=curDir, toolDir=toolDir)
        assert str(configPath) == str(expected)
    finally:
        deletelocalconfig()


def test_getconfigpath_intool():
    curDir = Path.cwd()
    toolDir = gettooldir()
    createtoolconfig()
    try:
        expected = toolDir.joinpath("config.json")
        configPath = main.getconfigpath(currentDir=curDir, toolDir=toolDir)
        assert str(configPath) == str(expected)
    finally:
        deletetoolconfig()
