import sys
import json
from pathlib import Path

with open("config.json", encoding="utf-8") as fc:
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
