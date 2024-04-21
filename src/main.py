import sys
import os
import json
from pathlib import Path
from datetime import datetime
from util.types import Config


def printstartmessage(currentDir: Path, toolDir: Path) -> None:
    print("****************************************")
    print(f"VBAUnit kicked on {currentDir}")
    print(f"Tool is in {toolDir}")
    print(f"{datetime.now()}")
    print()


def changecurdir(newDir: str) -> None:
    runDir = Path(newDir).resolve()
    os.chdir(runDir)


def getconfigpath(currentDir: Path, toolDir: Path) -> Path:
    configPath = currentDir.joinpath("config.json")
    if not configPath.exists():
        print("[warn] config get from tool directory")
        configPath = toolDir.joinpath("config.json")

    return configPath


def getconfig(configPath: Path) -> Config:
    with open(str(configPath), encoding="utf-8") as fc:
        jsondict = json.load(fc)

    config = Config()
    config.parse(jsondict=jsondict)
    return config


def getscenariopath(scenario: str) -> Path:
    scenarioPath = Path(config.scenario)
    if scenarioPath.is_absolute():
        print("absolute path")
    else:
        print(f"relative path based on {Path.cwd()}")
        scenarioPath = scenarioPath.resolve()
        print(f"resolved: {scenarioPath}")

    return scenarioPath


def getbridgepath(toolDir: Path) -> Path:
    return toolDir.joinpath("VBAUnitCOMBridge.xlsm")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 引数があればカレントディレクトリに設定
        changecurdir(sys.argv[1])

    currentDir = Path.cwd()
    toolDir = Path(__file__).parent
    printstartmessage(currentDir=currentDir, toolDir=toolDir)

    cocnfigPath = getconfigpath(currentDir=currentDir, toolDir=toolDir)
    if not cocnfigPath.exists():
        print("config.json not exists")
        sys.exit()

    config = getconfig(configPath=cocnfigPath)
    print(f"start with: {config.scenario}")

    scenarioPath = getscenariopath(config.scenario)
    if not scenarioPath.exists():
        print("scenario not exists")
        sys.exit()

    print()
    print("run test!")

    bridgePath = getbridgepath(toolDir=toolDir)

    print(f"using: {bridgePath}")
