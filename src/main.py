import sys
import os
from pathlib import Path
from datetime import datetime
from util.types import TestScenario, TestSuite, TestScope
from runner.run import run_testsuite
from vbaunit_lib.testlib import setglobalbridgepath


def __isvalidargv(argv: list[str]) -> bool:
    if len(argv) == 1:
        # 引数が最低1つ必要
        return False
    if len(argv) % 2 == 0:
        # 必ず奇数個
        return False
    return True


def __getargvalue(argv: list[str], key: str) -> str | None:
    if key in argv:
        keyindex = argv.index(key)
        value = argv[keyindex + 1]
        del argv[keyindex]  # keyを削除
        del argv[keyindex]  # valueを削除
        return value
    return None


def initenv(argv: list[str]) -> dict[str, str | Path]:
    """コマンドライン引数の解析と環境設定
    引数の順序
    1st: テストシナリオ絶対パス (必須)
    2nd: -w 作業ディレクトリ (任意)
    3rd: -o 出力先フォルダ相対パス（任意）
    4th: -n テストスイート名（任意）
    5th: -d テストスイートの説明（任意）
    6th: -f 実行フィルタ（任意）
    7th: -i 無視フィルタ（任意）
    8th: -s 全部か最後に失敗したものだけか（任意）
    """
    if not __isvalidargv(argv):
        print("Usage:")
        print(
            "python main.py {testsuite path} [-w {working directory}][-o {output directory}][-n {test suite name}]"
            "[-d {test suite description}][-f {execution filter}][-i {ignore filter}][-s {scope}]"
        )
        print()
        print("testscenario path: absolute path for scenario file (required)")
        print(
            "working directory: relational path for output directory based on scenario folder (optional)"
        )
        print(
            "output directory: relational path for output directory based on scenario folder (optional)"
        )
        print("test suite name: name for the test suite (optional)")
        print("test suite description: subject for the test suite (optional)")
        print("execution filter: filter for the tests to be executed (optional)")
        print("ignore filter: filter for the tests to be ignored (optional)")
        print("scope: all or lastfailed (optional)")
        print()
        sys.exit()

    argstack = argv.copy()

    scenariopath = Path(argv[1]).resolve()
    programdir = Path(__file__).parent.resolve()

    args: dict[str, str | Path] = {
        "scenario": scenariopath,
        "work": programdir,
        "out": scenariopath.parent.joinpath("results"),
        "name": f"VBAUnit {datetime.now():%Y%m%d%H%M%S}",
        "subject": f"VBAUnit Test {scenariopath.name}",
        "filters": "",
        "ignores": "",
        "scope": "all",
    }

    try:
        workdir = __getargvalue(argstack, "-w")
        if workdir is not None:
            args["work"] = Path(workdir).resolve()
            changecurdir(workdir)

        outdir = __getargvalue(argstack, "-o")
        if outdir is not None:
            args["out"] = Path(outdir).resolve()

        name = __getargvalue(argstack, "-n")
        if name is not None:
            args["name"] = name

        subject = __getargvalue(argstack, "-j")
        if subject is not None:
            subject = subject.strip('"')  # ダブルクオートを外す
            args["subject"] = subject

        filters = __getargvalue(argstack, "-f")
        if filters is not None:
            args["filters"] = filters

        ignores = __getargvalue(argstack, "-i")
        if ignores is not None:
            args["ignores"] = ignores

        scope = __getargvalue(argstack, "-c")
        if scope is not None:
            if scope == "all" or scope == "lastfailed":
                args["scope"] = scope
    except IndexError as e:
        print(f"[ERR]Invalid argument format: {e}")
        # とりあえず継続する

    return args


def printstartmessage(currentdir: Path, tooldir: Path) -> None:
    print("****************************************")
    print(f"VBAUnit kicked on {currentdir}")
    print(f"Tool is in {tooldir}")
    print(f"{datetime.now()}")
    print()


def changecurdir(newdir: str) -> None:
    rundir = Path(newdir).resolve()
    os.chdir(rundir)


def getscenariopath(scenario: str) -> Path:
    scenariopath = Path(scenario)
    if scenariopath.is_absolute():
        print(f"absolute scenario path {scenariopath}")
    else:
        print(f"relative scenario path based on {Path.cwd()}")
        scenariopath = scenariopath.resolve()
        print(f"resolved: {scenariopath}")

    return scenariopath


def getbridgepath(tooldir: Path) -> Path:
    return tooldir.joinpath("VBAUnitCOMBridge.xlsm")


if __name__ == "__main__":
    testconfig = initenv(sys.argv)

    # 開始メッセージ

    currentdir = Path.cwd()
    tooldir = Path(__file__).parent
    printstartmessage(currentdir=currentdir, tooldir=tooldir)

    # テストパス設定

    if not Path(testconfig["scenario"]).exists():
        print("scenario not exists")
        sys.exit()

    bridgepath = getbridgepath(tooldir=tooldir)
    setglobalbridgepath(bridgepath)

    print(f"using bridge: {bridgepath}")

    sys.path.append(str(tooldir))  # テストコードの方でvbaunit_libが使えるようになる

    # テストスイート

    scenario = TestScenario(Path(str(testconfig["scenario"])))

    testsuite = TestSuite(
        name=str(testconfig["name"]),
        subject=str(testconfig["subject"]),
        scenario=scenario,
        filters=str(testconfig["filters"]),
        ignores=str(testconfig["ignores"]),
        scope=TestScope(testconfig["scope"]),
    )

    # テスト実行

    print()
    print("run test!")

    run_testsuite(
        testsuite,
        Path(str(testconfig["scenario"])),
        bridgepath,
        Path(testconfig["out"]),
    )
