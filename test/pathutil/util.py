from pathlib import Path

def gettooldir() -> Path:
    return Path(__file__).parent.parent.parent.joinpath("src")
