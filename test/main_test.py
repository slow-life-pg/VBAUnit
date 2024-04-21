from pathlib import Path
from pathutil.util import gettooldir
from configutil.util import (
    ConfigPlace,
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


def getconfigpath_boilerplate(place: ConfigPlace) -> None:
    curDir = Path.cwd()
    toolDir = gettooldir()
    filePath = createconfigfile(place=place)
    try:
        configPath = main.getconfigpath(currentDir=curDir, toolDir=toolDir)
        assert str(configPath) == str(filePath)
    finally:
        deleteconfigfile(place=place)


def test_getconfigpath_incurrent():
    getconfigpath_boilerplate(ConfigPlace.LOCAL)


def test_getconfigpath_intool():
    getconfigpath_boilerplate(ConfigPlace.TOOL)


def test_getconfigpath_inoutside():
    outsidePath = Path.cwd().joinpath("outside")
    if not outsidePath.exists():
        outsidePath.mkdir()
    main.changecurdir(str(outsidePath))
    getconfigpath_boilerplate(ConfigPlace.OUTSIDE)
