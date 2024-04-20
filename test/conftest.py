import sys
from pathlib import Path

srcDir = Path(__file__).parent.parent.joinpath("src")
sys.path.append(str(srcDir))
