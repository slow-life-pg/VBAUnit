from pathlib import Path
import xlwings as xl

__bridgepath = Path(__file__).parent


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

    def exitapp(self) -> None:
        self.closeexcel()
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
        if self.__book:
            self.__book.api.VBProject.References.AddFromFile(excelpath)

        return self.__book

    def closeexcel(self) -> None:
        if self.__book is not None:
            self.__book.close()
            self.__book = None

    def getregexobj(self) -> object:
        """bridgeからRegexを取得"""
        pass

    def getcollectionobj(self) -> object:
        """bridgeからCollectionを取得"""
        pass

    def getdictionaryobj(self) -> object:
        """bridgeからDictionaryを取得"""
        pass

    def freeobj(self, obj: object) -> None:
        """bridgeから取得したオブジェクトを解放"""
        pass

    def callmacro(self, obj: object, macro_name: str, *args) -> object:
        """bridgeからマクロを呼び出す"""
        if self.__book:
            if len(args) <= 16:
                vbamacro = self.__book.macro(macro_name)
                res: list[object] = vbamacro(*args)
                return res[0] if res else None
            else:
                print(f"callmacro: too many macro arguments {len(args)}")
                return None
        else:
            print("callmacro: no book has opened")
            return None


def gettestlib(withapp: bool = False) -> VBAUnitTestLib:
    return VBAUnitTestLib(withapp=withapp)
