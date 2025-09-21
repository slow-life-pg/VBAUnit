import inspect
from functools import wraps
from pathlib import Path
from contextlib import contextmanager
import xlwings as xl

__globalbridgepath = Path(__file__).parent


def expect(requirement: bool, msg: str = "") -> None:
    """期待値を検証する関数。requirementがTrueであれば成功、Falseであれば失敗とする。"""
    if requirement:
        return

    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        raise AssertionError("Could not inspect caller frame")
    caller = frame.f_back
    info = inspect.getframeinfo(caller)
    # その行のソースコードを取得（存在する場合）
    src_line = info.code_context[0].strip() if info.code_context else "<source unavailable>"
    raise AssertionError(f"Check failed at {info.filename}:{info.lineno}\n  -> {src_line}" + (f"\n  msg: {msg}" if msg else ""))


def description(subject: str):
    """テストケースの説明を付与するデコレーター。"""

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)

        wrapper._subject = subject  # type: ignore
        return wrapper

    return decorator


def ignore(func):
    """無視したいテストケースを識別するデコレーター。"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)

    wrapper._is_ignored = True  # type: ignore
    return wrapper


def setglobalbridgepath(bridgepath: Path) -> None:
    global __globalbridgepath
    __globalbridgepath = bridgepath


def getglobalbridgepath() -> Path:
    global __globalbridgepath
    return __globalbridgepath


class VBAUnitTestLib:
    def __init__(self, __globalbridgepath: Path, withapp: bool = True, visible: bool = False) -> None:
        self.__bridgepath = __globalbridgepath
        self.__withapp = withapp
        self.__visible = visible

        if withapp:
            self.__app = xl.App(visible=visible)
        else:
            self.__app = None
        self.__book = None
        self.__internalbook = None

        self.__comobjects: list[object] = []

    @property
    def appready(self) -> bool:
        return self.__app is not None

    @contextmanager
    def runapp(self, excelpath: str):
        try:
            book = self.openexcel(excelpath)
            yield book
        finally:
            self.exitapp()

    def exitapp(self) -> None:
        self.closeexcel()
        self.__closeapp()

    def __closeapp(self) -> None:
        if self.__app is not None:
            self.__app.quit()
            self.__app = None

    def openexcel(self, excelpath: str) -> xl.Book:
        # Excelファイルを開く処理を実装
        if self.__app is None:
            self.__app = xl.App(visible=self.__visible)
        self.__app.display_alerts = False
        # self.__app.screen_updating = False

        self.__book = self.__app.books.open(self.__bridgepath, update_links=True, ignore_read_only_recommended=True)
        # 開けなかったらErrorが出ているはずだから独自にthrowしない。
        if self.__book:
            excelfullpath = Path(excelpath).resolve()
            # self.__book.api.VBProject.References.AddFromFile(excelfullpath)
            self.__internalbook = self.__app.books.open(excelfullpath)

        return self.__internalbook

    def closeexcel(self) -> None:
        if self.__internalbook is not None:
            self.__internalbook.close()
            self.__internalbook = None
        if self.__book is not None:
            self.__book.close()
            self.__book = None
            self.freeobjs()
        if self.__withapp:
            self.__closeapp()

    def getregexobj(self) -> object:
        """bridgeからRegexを取得"""
        if self.__book:
            regexobj = self.__app.api.Run(self.__getbridgemacroname("GetRegexp"))
            self.__comobjects.append(regexobj)
            return regexobj
        else:
            return None

    def getcollectionobj(self) -> object:
        """bridgeからCollectionを取得"""
        if self.__book:
            collectionobj = self.__app.api.Run(self.__getbridgemacroname("GetNewCollection"))
            self.__comobjects.append(collectionobj)
            return collectionobj
        else:
            return None

    def getdictionaryobj(self) -> object:
        """bridgeからDictionaryを取得"""
        if self.__book:
            dictionaryobj = self.__app.api.Run(self.__getbridgemacroname("GetNewDictionary"))
            self.__comobjects.append(dictionaryobj)
            return dictionaryobj
        else:
            return None

    def freeobj(self, obj: object) -> None:
        """bridgeから取得したオブジェクトを解放"""
        if self.__book:
            if obj:
                self.__app.api.Run(self.__getbridgemacroname("Free", obj))
            if obj in self.__comobjects:
                self.__comobjects.remove(obj)

    def freeobjs(self) -> None:
        """bridgeから取得した全てのオブジェクトを解放"""
        if self.__book:
            freemacro = self.__book.macro("Free")
            for obj in self.__comobjects:
                freemacro(obj)
            self.__comobjects.clear()

    def callmacro(self, obj: object, macro_name: str, *args) -> list[object]:
        return self.__callmacro(obj, False, macro_name, *args)

    def callcreativemacro(self, obj: object, macro_name: str, *args) -> list[object]:
        res = self.__callmacro(obj, True, macro_name, *args)
        if res[0]:
            self.__comobjects.append(res[0])
        return res

    def registercomobject(self, obj: object):
        if obj:
            self.__comobjects.append(obj)

    def __callmacro(self, obj: object, creation: bool, macro_name: str, *args) -> list[object]:
        """bridgeからマクロを呼び出す。"""
        if self.__book:
            if len(args) <= 16:
                vbamacro = self.__book.macro("CallMacro")
                vbargs = list(args)
                res: list[object] = vbamacro(obj, creation, self.__internalbook, macro_name, vbargs)
                self.__comobjects.append(res)
                return res
            else:
                print(f"callmacro: too many macro arguments {len(args)}. Must be <= 16.")
                return [None]
        else:
            print("callmacro: no book has opened")
            return [None]

    def __getbridgemacroname(self, macro: str) -> str:
        return "VBAUnitCOMBridge.xlsm!" + macro


def gettestlib(withapp: bool = False, visible: bool = False) -> VBAUnitTestLib:
    return VBAUnitTestLib(__globalbridgepath, withapp=withapp, visible=visible)
