if __name__ == "__main__":
    import sys
    import json
    from pathlib import Path
    from datetime import datetime

    toolDir = Path(__file__).parent
    print("****************************************")
    print(f"VBAUnit kicked on {Path.cwd()}")
    print(f"Tool is in {toolDir}")
    print(f"{datetime.now()}")
    print()

    configPath = toolDir.joinpath("config.json")

    with open(str(configPath), encoding="utf-8") as fc:
        config = json.load(fc)

    print(f"start with: {config['scenario']}")

    scenarioPath = Path(config["scenario"])
    if scenarioPath.is_absolute():
        print("absolute path")
    else:
        print(f"relative path based on {Path.cwd()}")
        scenarioPath = scenarioPath.resolve()
        print(f"resolved: {scenarioPath}")

    if not scenarioPath.exists():
        print("scenario not exists")
        sys.exit()

    print()
    print("run test!")

    bridgePath = toolDir.joinpath("Bridge.xlsm")

    print(f"using: {bridgePath}")
