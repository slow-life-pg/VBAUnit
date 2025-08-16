import sys
import os
import json
from pathlib import Path
from datetime import datetime
from util.types import Config
from runner.run import run_testsuite


def analyzeargs(argv: list[str]) -> str:
    if len(argv) == 1 or len(argv) > 3:
        print("Usage:")
        print("python main.py {testsuite name} [{working directory}]")
        print()
        sys.exit()

    if len(argv) > 2:
        # 引数があればカレントディレクトリに設定
        argpath = Path(argv[2])
        if argpath.exists():
            changecurdir(argpath)

    return argv[1]


def printstartmessage(currentdir: Path, tooldir: Path) -> None:
    print("****************************************")
    print(f"VBAUnit kicked on {currentdir}")
    print(f"Tool is in {tooldir}")
    print(f"{datetime.now()}")
    print()


def changecurdir(newdir: Path) -> None:
    rundir = newdir.resolve()
    os.chdir(rundir)


def getconfigpath(currentdir: Path, tooldir: Path) -> Path:
    configpath = currentdir.joinpath("config.json")
    if not configpath.exists():
        print("[warn] config get from tool directory")
        configpath = tooldir.joinpath("config.json")

    return configpath


def getconfig(configpath: Path) -> Config:
    with open(str(configpath), encoding="utf-8") as fc:
        jsondict = json.load(fc)

    testconfig = Config(jsondict=jsondict)
    return testconfig


def getscenariopath(scenario: str) -> Path:
    scenariopath = Path(scenario)
    if scenariopath.is_absolute():
        print("absolute scenario path")
    else:
        print(f"relative scenario path based on {Path.cwd()}")
        scenariopath = scenariopath.resolve()
        print(f"resolved: {scenariopath}")

    return scenariopath


def getbridgepath(tooldir: Path) -> Path:
    return tooldir.joinpath("VBAUnitCOMBridge.xlsm")


if __name__ == "__main__":
    testsuite = analyzeargs(sys.argv)

    # Config取得

    currentdir = Path.cwd()
    tooldir = Path(__file__).parent
    printstartmessage(currentdir=currentdir, tooldir=tooldir)

    cocnfigpath = getconfigpath(currentdir=currentdir, tooldir=tooldir)
    if not cocnfigpath.exists():
        print("config.json not exists")
        sys.exit()

    config = getconfig(configpath=cocnfigpath)
    print(f"start with: {config.scenario}")

    # テストパス設定

    scenariopath = getscenariopath(config.scenario)
    if not scenariopath.exists():
        print("scenario not exists")
        sys.exit()

    bridgepath = getbridgepath(tooldir=tooldir)

    print(f"using: {bridgepath}")

    sys.path.append(str(tooldir))  # テストコードの方でvbaunit_libが使えるようになる

    # テスト実行

    print()
    print("run test!")

    run_testsuite(config=config)
