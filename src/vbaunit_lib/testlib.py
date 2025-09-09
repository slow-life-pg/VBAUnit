from functools import wraps
from pathlib import Path
import xlwings as xl

__bridgepath = Path(__file__).parent


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


def setbridgepath(bridgepath: Path) -> None:
    global __bridgepath
    __bridgepath = bridgepath


class VBAUnitTestLib:
    def __init__(self, withapp: bool = True) -> None:
        self.__bridgepath = __bridgepath
        self.__withapp = withapp
        if withapp:
            self.__app = xl.App(visible=False)
        else:
            self.__app = None
        self.__book = None
        self.__internalbookname = ""

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
            self.__app = xl.App(visible=False)
        self.__app.display_alerts = False
        self.__app.screen_updating = False

        self.__book = self.__app.books.open(self.__bridgepath)
        # 開けなかったらErrorが出ているはずだから独自にthrowしない。
        if self.__book:
            self.__book.api.VBProject.References.AddFromFile(excelpath)
            self.__internalbookname = Path(excelpath).name

        return self.__book

    def closeexcel(self) -> None:
        if self.__book is not None:
            self.__book.close()
            self.__book = None
            if self.__withapp:
                self.__closeapp()

    def getregexobj(self) -> object:
        """bridgeからRegexを取得"""
        if self.__book:
            vbamacro = self.__book.macro("GetRegexp")
            regexobj = vbamacro()
            return regexobj
        else:
            return None

    def getcollectionobj(self) -> object:
        """bridgeからCollectionを取得"""
        if self.__book:
            vbamacro = self.__book.macro("GetNewCollection")
            collectionobj = vbamacro()
            return collectionobj
        else:
            return None

    def getdictionaryobj(self) -> object:
        """bridgeからDictionaryを取得"""
        if self.__book:
            vbamacro = self.__book.macro("GetNewDictionary")
            dictionaryobj = vbamacro()
            return dictionaryobj
        else:
            return None

    def freeobj(self, obj: object) -> None:
        """bridgeから取得したオブジェクトを解放"""
        if self.__book:
            vbamacro = self.__book.macro("Free")
            vbamacro(obj)

    def callmacro(self, obj: object, macro_name: str, *args) -> object:
        """bridgeからマクロを呼び出す"""
        if self.__book:
            if len(args) <= 16:
                vbamacro = self.__book.macro(macro_name)
                res: list[object] = vbamacro(
                    obj, self.__internalbookname, macro_name, *args
                )
                return res[0] if res else None
            else:
                print(
                    f"callmacro: too many macro arguments {len(args)}. Must be <= 16."
                )
                return None
        else:
            print("callmacro: no book has opened")
            return None


def gettestlib(withapp: bool = False) -> VBAUnitTestLib:
    return VBAUnitTestLib(withapp=withapp)
