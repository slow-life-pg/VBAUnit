import sys
import os
from pathlib import Path
from datetime import datetime
from util.types import TestScenario, TestSuite, TestScope
from runner.run import run_testsuite
from vbaunit_lib.testlib import setglobalbridgepath


def __isvalidargv(argv: list[str]) -> bool:
    # 先頭要素は"main.py"
    if len(argv) == 2:
        # 引数が最低1つ必要
        return True
    if len(argv) % 2 == 0:
        # 必ず奇数個
        return True
    return False


def __getargvalue(argv: list[str], key: str) -> str | None:
    if key in argv:
        keyindex = argv.index(key)
        value = argv[keyindex + 1]
        del argv[keyindex]  # keyを削除
        del argv[keyindex]  # valueを削除
        return value
    return None


def __createoutpudirectory(dir: Path) -> None:
    if str(dir) == "" or dir.exists():
        return
    __createoutpudirectory(dir.parent)
    dir.mkdir()


def initenv(argv: list[str]) -> dict[str, str | Path]:
    """コマンドライン引数の解析と環境設定
    引数の順序
    1st: テストシナリオ絶対パス (必須)
    2nd: -w 作業ディレクトリ (任意)
    3rd: -o 出力先フォルダ相対パス（任意）
    4th: -n テストスイート名（任意）
    5th: -d テストスイートの説明（任意）
    6th: -k 実行フィルタ（任意）
    7th: -i 無視フィルタ（任意）
    8th: -s 全部か最後に失敗したものだけか（任意）
    """
    if not __isvalidargv(argv):
        print("Usage:")
        print(
            "python main.py {testsuite path} [-w {working directory}][-o {output directory}][-n {test suite name}]"
            "[-d {test suite description}][-g {group name}][-k {execution filter}][-i {ignore filter}][-s {scope}]"
        )
        print()
        print("testscenario path: absolute path for scenario file (required)")
        print("working directory: relational path for output directory based on scenario folder (optional)")
        print("output directory: relational path for output directory based on scenario folder (optional)")
        print("test suite name: name for the test suite (optional)")
        print("test suite description: subject for the test suite (optional)")
        print("execution group name: group name for the tests to be executed (optional)")
        print("execution filter: filter for the tests to be executed (optional)")
        print("ignore filter: filter for the tests to be ignored (optional)")
        print("scope: all or lastfailed (optional)")
        print()
        sys.exit()

    argstack = argv.copy()
    argstack.pop(0)  # 先頭を除去

    patharg = argstack.pop(0)
    scenariopath = Path(patharg).resolve()
    workingdir = scenariopath.parent

    args: dict[str, str | Path] = {
        "scenario": scenariopath,
        "work": workingdir,
        "out": workingdir.joinpath("results"),
        "name": f"VBAUnit {datetime.now():%Y%m%d%H%M%S}",
        "subject": f"VBAUnit Test {scenariopath.name}",
        "groups": "",
        "filters": "",
        "ignores": "",
        "scope": "all",
    }

    try:
        workdir = __getargvalue(argstack, "-w")
        if workdir is not None:
            args["work"] = Path(workdir).resolve()

        outdir = __getargvalue(argstack, "-o")
        if outdir is not None:
            args["out"] = Path(args["work"]).joinpath(Path(outdir))

        name = __getargvalue(argstack, "-n")
        if name is not None:
            args["name"] = name

        subject = __getargvalue(argstack, "-j")
        if subject is not None:
            subject = subject.strip('"')  # ダブルクオートを外す
            args["subject"] = subject

        groups = __getargvalue(argstack, "-g")
        if groups is not None:
            args["groups"] = groups

        filters = __getargvalue(argstack, "-k")
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
    print(f"{datetime.now():%Y/%m/%d %H:%M:%S}")
    print()


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
        print(f"scenario not exists {testconfig['scenario']}")
        sys.exit()

    bridgepath = getbridgepath(tooldir=tooldir)
    setglobalbridgepath(bridgepath)

    print(f"using bridge: {bridgepath}")

    sys.path.append(str(tooldir.joinpath("src")))  # テストコードの方でvbaunit_libが使えるようになる

    __createoutpudirectory(Path(testconfig["out"]))

    # テストスイート

    os.chdir(testconfig["work"])

    scenario = TestScenario(Path(str(testconfig["scenario"])))

    testsuite = TestSuite(
        name=str(testconfig["name"]),
        subject=str(testconfig["subject"]),
        scenario=scenario,
        groups=str(testconfig["groups"]),
        filters=str(testconfig["filters"]),
        ignores=str(testconfig["ignores"]),
        scope=TestScope(testconfig["scope"]),
    )

    # テスト実行

    print()
    print(f"{testsuite.count} tests will be executed.")
    print("run test!")

    run_testsuite(
        testsuite,
        Path(str(testconfig["scenario"])),
        bridgepath,
        Path(testconfig["out"]),
    )

    os.chdir(currentdir)
