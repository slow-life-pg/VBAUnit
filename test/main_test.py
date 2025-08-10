from pathlib import Path
from pathutil.util import gettooldir
from configutil.util import (
    ConfigPlace,
    getlocalscenariofilenam,
    createconfigfile,
    deleteconfigfile,
)
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


def assertconfigpath_boilerplate(place: ConfigPlace) -> None:
    curdir = Path.cwd()
    tooldir = gettooldir()
    filepath = createconfigfile(place=place)
    try:
        configpath = main.getconfigpath(currentdir=curdir, tooldir=tooldir)
        assert str(configpath) == str(filepath)
    finally:
        deleteconfigfile(place=place)


def test_getconfigpath_incurrent():
    assertconfigpath_boilerplate(ConfigPlace.LOCAL)


def test_getconfigpath_intool():
    assertconfigpath_boilerplate(ConfigPlace.TOOL)


def test_getconfigpath_inoutside():
    outsidepath = Path.cwd().joinpath("outside")
    if not outsidepath.exists():
        outsidepath.mkdir()
    main.changecurdir(outsidepath)
    assertconfigpath_boilerplate(ConfigPlace.OUTSIDE)


def test_getconfig_normal():
    configpath = createconfigfile(ConfigPlace.LOCAL)
    config = main.getconfig(configpath=configpath)
    assert config.scenario == getlocalscenariofilenam()


def test_getscenariopath_relative():
    cwd_scenariopath = Path.cwd().joinpath(getlocalscenariofilenam())
    assert main.getscenariopath(getlocalscenariofilenam()) == cwd_scenariopath


def test_getscenariopath_absolute():
    abs_scenariopath = Path("c:\\test").joinpath(getlocalscenariofilenam())
    assert main.getscenariopath(str(abs_scenariopath)) == abs_scenariopath
