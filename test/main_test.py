from pathlib import Path
from pathutil.util import gettooldir
from configutil.util import (
    ConfigPlace,
    getlocalscenariofilenam,
    createconfigfile,
    deleteconfigfile,
)
import main

#### このやり方もあるにはある　→　今回はconftest.pyで行ってみる
# import pytest
# @pytest.fixture
# def main_fixture():
#     # ここで検索パスを追加して
#     srcDir = Path(__file__).parent.parent.joinpath("src")
#     sys.path.append(str(srcDir))

#     # ここでimportしてもテストケースの実行時にModuleNotFoundError

#     yield

#     sys.path.remove(str(srcDir))


def assertconfigpath_boilerplate(place: ConfigPlace) -> None:
    curDir = Path.cwd()
    toolDir = gettooldir()
    filePath = createconfigfile(place=place)
    try:
        configPath = main.getconfigpath(currentDir=curDir, toolDir=toolDir)
        assert str(configPath) == str(filePath)
    finally:
        deleteconfigfile(place=place)


def test_getconfigpath_incurrent():
    assertconfigpath_boilerplate(ConfigPlace.LOCAL)


def test_getconfigpath_intool():
    assertconfigpath_boilerplate(ConfigPlace.TOOL)


def test_getconfigpath_inoutside():
    outsidePath = Path.cwd().joinpath("outside")
    if not outsidePath.exists():
        outsidePath.mkdir()
    main.changecurdir(outsidePath)
    assertconfigpath_boilerplate(ConfigPlace.OUTSIDE)


def test_getconfig_normal():
    configPath = createconfigfile(ConfigPlace.LOCAL)
    config = main.getconfig(configPath=configPath)
    assert config.scenario == getlocalscenariofilenam()


def test_getscenariopath_relative():
    cwdScenarioPath = Path.cwd().joinpath(getlocalscenariofilenam())
    assert main.getscenariopath(getlocalscenariofilenam()) == cwdScenarioPath


def test_getscenariopath_absolute():
    absScenarioPath = Path("c:\\test").joinpath(getlocalscenariofilenam())
    assert main.getscenariopath(str(absScenarioPath)) == absScenarioPath
