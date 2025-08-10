import sys
from pathlib import Path

srcdir = Path(__file__).parent.parent.joinpath("src")
sys.path.append(str(srcdir))
