import inspect
import psutil
import gc
from functools import wraps
from pathlib import Path
from contextlib import contextmanager
from win32com.client import VARIANT
import pythoncom
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


def expect_collection(collectionobj: object, key: str, expectation: object, msg: str = "") -> None:
    """Collectionオブジェクトのキーに対応する期待値を検証する関数。keyから値を取得してexpectationと比較する。"""
    variantkey = VARIANT(pythoncom.VT_BSTR, key)
    valueofkey = collectionobj.Item(variantkey)  # type: ignore
    if valueofkey == expectation:
        return

    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        raise AssertionError("Could not inspect caller frame")
    caller = frame.f_back
    info = inspect.getframeinfo(caller)
    # その行のソースコードを取得（存在する場合）
    src_line = info.code_context[0].strip() if info.code_context else "<source unavailable>"
    raise AssertionError(f"Check failed at {info.filename}:{info.lineno}\n  -> {src_line}" + (f"\n  msg: {msg}" if msg else ""))


def expect_collection_list(collectionobj: object, expectation: list[object], msg: str = "") -> None:
    """Collectionオブジェクトとexpectationの全要素が順序を含めて一致することを確認する。"""
    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        raise AssertionError("Could not inspect caller frame")
    caller = frame.f_back
    info = inspect.getframeinfo(caller)
    # その行のソースコードを取得（存在する場合）
    src_line = info.code_context[0].strip() if info.code_context else "<source unavailable>"

    if collectionobj.Count() != len(expectation):  # type: ignore
        raise AssertionError(
            f"Check failed: collection count {collectionobj.Count()} != expectation count {len(expectation)}"  # type: ignore
            + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
            + (f"\n  msg: {msg}" if msg else "")
        )
    for i in range(1, collectionobj.Count() + 1):  # type: ignore
        valueofkey = collectionobj.Item(i)  # type: ignore
        if valueofkey != expectation[i - 1]:
            raise AssertionError(
                f"Check failed: collection item {i} value {valueofkey} != expectation value {expectation[i - 1]}"  # type: ignore
                + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
                + (f"\n  msg: {msg}" if msg else "")
            )
    # どこにも筆禍からなければ成功
    return


def expect_collection_dict(collectionobj: object, expectation: dict[str, object], msg: str = "") -> None:
    """Collectionオブジェクトとexpectationのキーと値のペアが、一致することを確認する。"""
    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        raise AssertionError("Could not inspect caller frame")
    caller = frame.f_back
    info = inspect.getframeinfo(caller)
    # その行のソースコードを取得（存在する場合）
    src_line = info.code_context[0].strip() if info.code_context else "<source unavailable>"

    if collectionobj.Count() != len(expectation):  # type: ignore
        raise AssertionError(
            f"Check failed: collection count {collectionobj.Count()} != expectation count {len(expectation)}"  # type: ignore
            + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
            + (f"\n  msg: {msg}" if msg else "")
        )
    for key in expectation.keys():
        valueofkey = collectionobj.Item(key)  # type: ignore
        if valueofkey != expectation[key]:
            raise AssertionError(
                f"Check failed: collection item {key} value {valueofkey} != expectation value {expectation[key]}"  # type: ignore
                + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
                + (f"\n  msg: {msg}" if msg else "")
            )
    # どこにも筆禍からなければ成功
    return


def expect_dictionary(dictionaryobj: object, expectation: dict[object, object], msg: str = "") -> None:
    """Dictionaryオブジェクトとexpectationのキーと値のペアが、一致することを確認する。"""
    frame = inspect.currentframe()
    if frame is None or frame.f_back is None:
        raise AssertionError("Could not inspect caller frame")
    caller = frame.f_back
    info = inspect.getframeinfo(caller)
    # その行のソースコードを取得（存在する場合）
    src_line = info.code_context[0].strip() if info.code_context else "<source unavailable>"

    if dictionaryobj.Count != len(expectation):  # type: ignore
        raise AssertionError(
            f"Check failed: dictionary count {dictionaryobj.Count} != expectation count {len(expectation)}"  # type: ignore
            + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
            + (f"\n  msg: {msg}" if msg else "")
        )
    for key in expectation.keys():
        valueofkey = dictionaryobj.Item(key)  # type: ignore
        if valueofkey != expectation[key]:
            raise AssertionError(
                f"Check failed: dictionary item {key} value {valueofkey} != expectation value {expectation[key]}"  # type: ignore
                + f" at {info.filename}:{info.lineno}\n  -> {src_line}"
                + (f"\n  msg: {msg}" if msg else "")
            )
    # どこにも筆禍からなければ成功
    return


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
    __VBEXT_CT_STDMODULE = 1  # 標準モジュール
    __VBEXT_CT_CLASSMODULE = 2  # クラスモジュール

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

        self.__pid = -1

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
        try:
            self.freeobjs()
        finally:
            self.closeexcel()
            self.__closeapp()

    def __closeapp(self) -> None:
        if self.__app is not None:
            self.__app.quit()
            self.__app = None
            self.__killserver()

    def __killserver(self) -> None:
        if self.__pid > 0:
            try:
                p = psutil.Process(self.__pid)
            except psutil.NoSuchProcess:
                self.__pid = -1
                return
            try:
                p.terminate()
                p.wait(3)
            except Exception as e:
                print(f"Could not kill Excel process {self.__pid}: {e}")
            finally:
                self.__pid = -1

    def openexcel(self, excelpath: str) -> xl.Book:
        # Excelファイルを開く処理を実装
        if self.__app is None:
            self.__app = xl.App(visible=self.__visible)
        self.__pid = self.__app.pid

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
                self.__app.api.Run(self.__getbridgemacroname("Free"), obj)
            if obj in self.__comobjects:
                self.__comobjects.remove(obj)
            obj = None

    def freeobjs(self) -> None:
        """bridgeから取得した全てのオブジェクトを解放"""
        if self.__book:
            freemacro = self.__book.macro("Free")
            for obj in self.__comobjects:
                # 全削除の時は既に解放済みかもしれない
                try:
                    freemacro(obj)
                except Exception as e:
                    print(f"freeobjs error: {e}")
                    print("  skip this object")
                obj = None
            self.__comobjects.clear()
            gc.collect()

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
                res: list[object] = vbamacro(obj, creation, self.__internalbook.name, macro_name, vbargs)
                # self.__comobjects.append(res) # 呼び出し結果はスコープを抜けると解放されるので登録しない
                return res
            else:
                print(f"callmacro: too many macro arguments {len(args)}. Must be <= 16.")
                return [None]
        else:
            print("callmacro: no book has opened")
            return [None]

    def __getbridgemacroname(self, macro: str) -> str:
        return "VBAUnitCOMBridge.xlsm!" + macro

    def create_newinstance(self, class_name: str) -> object:
        """bridgeから指定されたClassモジュールのインスタンスを取得"""
        if self.__book:
            temp_component = self.__internalbook.api.VBProject.VBComponents.Add(1)  # 1:標準モジュール

            tempcode = (
                f"Public Function GetInstanceOf__{class_name}() As Variant\n"
                + f"    Set GetInstanceOf__{class_name} = New {class_name}\n"
                + "End Function\n"
            )
            temp_component.CodeModule.AddFromString(tempcode)
            create_modulename = f"{temp_component.Name}.GetInstanceOf__{class_name}"
            createmacro = self.__internalbook.macro(create_modulename)
            instance = createmacro()
            self.registercomobject(instance)

            self.__internalbook.api.VBProject.VBComponents.Remove(temp_component)

            return instance
        else:
            return None

    def create_backdoor(self, module_name: str, code: str) -> None:
        """指定された標準モジュールまたはクラスモジュールに関数を追加する"""
        if self.__book:
            target_component = None
            for comp in self.__internalbook.api.VBProject.VBComponents:
                if comp.Name == module_name and comp.Type in (VBAUnitTestLib.__VBEXT_CT_STDMODULE, VBAUnitTestLib.__VBEXT_CT_CLASSMODULE):
                    target_component = comp
                    break
            if target_component is None:
                print(f"create_backdoor: module {module_name} not found")
                return

            code_module = target_component.CodeModule
            lastline = code_module.CountOfLines
            code_module.InsertLines(lastline + 1, code)


def gettestlib(withapp: bool = False, visible: bool = False) -> VBAUnitTestLib:
    return VBAUnitTestLib(__globalbridgepath, withapp=withapp, visible=visible)
