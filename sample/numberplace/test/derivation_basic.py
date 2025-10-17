from vbaunit_lib.testlib import gettestlib, expect, expect_equals  # noqa


def get_problemboard_beginner() -> list[list[int]]:
    problem = [[0 for _ in range(9)] for _ in range(9)]

    # puzzle for beginners
    # 5 . . | . 2 . | . . 8
    # . 3 . | 8 . . | 6 . .
    # . 4 8 | . . . | 1 . .
    # ------+-------+------
    # . . 6 | 3 . 7 | . 9 .
    # 3 . . | . . . | . . 7
    # . 8 . | 6 . 5 | 3 . .
    # ------+-------+------
    # . 6 . | . . . | 8 4 .
    # . . 3 | . . 8 | . 6 .
    # 9 . . | . 4 . | . . 3
    problem[0][0] = 5
    problem[0][4] = 2
    problem[0][8] = 8

    problem[1][1] = 3
    problem[1][3] = 8
    problem[1][6] = 6

    problem[2][1] = 4
    problem[2][2] = 8
    problem[2][6] = 1

    problem[3][2] = 6
    problem[3][3] = 3
    problem[3][5] = 7
    problem[3][7] = 9

    problem[4][0] = 3
    problem[4][8] = 7

    problem[5][1] = 8
    problem[5][3] = 6
    problem[5][5] = 5
    problem[5][6] = 3

    problem[6][1] = 6
    problem[6][6] = 8
    problem[6][7] = 4

    problem[7][2] = 3
    problem[7][5] = 8
    problem[7][7] = 6

    problem[8][0] = 9
    problem[8][4] = 4
    problem[8][8] = 3

    return problem


def test_derive_blockcandidates_1_1():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveBlockCandidates", problem, 0, 0, 0)
        expect(res[5] == 0)
        assert isinstance(res[4], int)
        result: int = res[4]  # blockcandidates argument ByRef
        expect(result & 0b0000_0000_0001 == 0b000000000001)  # 1 allowed
        expect(result & 0b0000_0000_0010 == 0b000000000010)  # 2 allowed
        expect(result & 0b0000_0000_0100 == 0)  # 3 unabled
        expect(result & 0b0000_0000_1000 == 0)  # 4 unabled
        expect(result & 0b0000_0001_0000 == 0)  # 5 unabled
        expect(result & 0b0000_0010_0000 == 0b000000100000)  # 6 allowed
        expect(result & 0b0000_0100_0000 == 0b000001000000)  # 7 allowed
        expect(result & 0b0000_1000_0000 == 0)  # 8 unabled
        expect(result & 0b0001_0000_0000 == 0b000100000000)  # 9 allowed


def test_derive_blockcandidates_2_3():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveBlockCandidates", problem, 3, 6, 0)
        expect(res[5] == 0)
        assert isinstance(res[4], int)
        result: int = res[4]  # blockcandidates argument ByRef
        expect(result & 0b0000_0000_0001 == 0b000000000001)  # 1 allowed
        expect(result & 0b0000_0000_0010 == 0b000000000010)  # 2 allowed
        expect(result & 0b0000_0000_0100 == 0)  # 3 unabled
        expect(result & 0b0000_0000_1000 == 0b000000001000)  # 4 allowed
        expect(result & 0b0000_0001_0000 == 0b000000010000)  # 5 allowed
        expect(result & 0b0000_0010_0000 == 0b000000100000)  # 6 allowed
        expect(result & 0b0000_0100_0000 == 0)  # 7 unabled
        expect(result & 0b0000_1000_0000 == 0b000010000000)  # 8 allowed
        expect(result & 0b0001_0000_0000 == 0)  # 9 unabled


def test_derive_blockscandidates():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveBlocksCandidates", problem, [0 for _ in range(9)])
        expect(res[3] == 0)
        assert isinstance(res[2], tuple)
        results: tuple[int, int, int, int, int, int, int, int, int] = res[2]  # blockscandidates argument ByRef
        expect(len(results) == 9)
        expect_equals(0b0001_0110_0011, results[0])
        expect_equals(0b0001_0111_1101, results[1])
        expect_equals(0b0001_0101_1110, results[2])
        expect_equals(0b0001_0101_1011, results[3])
        expect_equals(0b0001_1000_1011, results[4])
        expect_equals(0b0000_1011_1011, results[5])
        expect_equals(0b0000_1101_1011, results[6])
        expect_equals(0b0001_0111_0111, results[7])
        expect_equals(0b0001_0101_0011, results[8])


def test_derive_columncandidates_2():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveColumnCandidates", problem, 1, 0)
        expect(res[4] == 0)
        assert isinstance(res[3], int)
        result: int = res[3]  # blockcandidates argument ByRef
        expect_equals(0b000000000001, result & 0b0000_0000_0001)  # 1 allowed
        expect_equals(0b000000000010, result & 0b0000_0000_0010)  # 2 allowed
        expect_equals(0, result & 0b0000_0000_0100)  # 3 unabled
        expect_equals(0, result & 0b0000_0000_1000)  # 4 unabled
        expect_equals(0b000000010000, result & 0b0000_0001_0000)  # 5 allowed
        expect_equals(0, result & 0b0000_0010_0000)  # 6 unabled
        expect_equals(0b000001000000, result & 0b0000_0100_0000)  # 7 allowed
        expect_equals(0, result & 0b0000_1000_0000)  # 8 unabled
        expect_equals(0b000100000000, result & 0b0001_0000_0000)  # 9 allowed


def test_derive_columncandidates_7():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveColumnCandidates", problem, 6, 0)
        expect(res[4] == 0)
        assert isinstance(res[3], int)
        result: int = res[3]  # blockcandidates argument ByRef
        expect_equals(0, result & 0b0000_0000_0001)  # 1 unabled
        expect_equals(0b000000000010, result & 0b0000_0000_0010)  # 2 allowed
        expect_equals(0, result & 0b0000_0000_0100)  # 3 unabled
        expect_equals(0b000000001000, result & 0b0000_0000_1000)  # 4 allowed
        expect_equals(0b000000010000, result & 0b0000_0001_0000)  # 5 allowed
        expect_equals(0, result & 0b0000_0010_0000)  # 6 unabled
        expect_equals(0b000001000000, result & 0b0000_0100_0000)  # 7 allowed
        expect_equals(0, result & 0b0000_1000_0000)  # 8 unabled
        expect_equals(0b000100000000, result & 0b0001_0000_0000)  # 9 allowed


def test_derive_columnscandidates():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveColumnsCandidates", problem, [0 for _ in range(9)])
        expect(res[3] == 0)
        assert isinstance(res[2], tuple)
        results: tuple[int, int, int, int, int, int, int, int, int] = res[2]  # columnskscandidates argument ByRef
        expect(len(results) == 9)
        expect_equals(0b0000_1110_1011, results[0])
        expect_equals(0b0001_0101_0011, results[1])
        expect_equals(0b0001_0101_1011, results[2])
        expect_equals(0b0001_0101_1011, results[3])
        expect_equals(0b0001_1111_0101, results[4])
        expect_equals(0b0001_0010_1111, results[5])
        expect_equals(0b0001_0101_1010, results[6])
        expect_equals(0b0000_1101_0111, results[7])
        expect_equals(0b0001_0011_1011, results[8])


def test_derive_rowcandidates_4():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveRowCandidates", problem, 3, 0)
        expect(res[4] == 0)
        assert isinstance(res[3], int)
        result: int = res[3]  # blockcandidates argument ByRef
        expect_equals(0b000000000001, result & 0b0000_0000_0001)  # 1 allowed
        expect_equals(0b000000000010, result & 0b0000_0000_0010)  # 2 allowed
        expect_equals(0, result & 0b0000_0000_0100)  # 3 unabled
        expect_equals(0b000000001000, result & 0b0000_0000_1000)  # 4 allowed
        expect_equals(0b000000010000, result & 0b0000_0001_0000)  # 5 allowed
        expect_equals(0, result & 0b0000_0010_0000)  # 6 unabled
        expect_equals(0, result & 0b0000_0100_0000)  # 7 unabled
        expect_equals(0b000010000000, result & 0b0000_1000_0000)  # 8 allowed
        expect_equals(0, result & 0b0001_0000_0000)  # 9 unabled


def test_derive_rowscandidates():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveRowsCandidates", problem, [0 for _ in range(9)])
        expect(res[3] == 0)
        assert isinstance(res[2], tuple)
        results: tuple[int, int, int, int, int, int, int, int, int] = res[2]  # columnskscandidates argument ByRef
        expect(len(results) == 9)
        expect_equals(0b0001_0110_1101, results[0])
        expect_equals(0b0001_0101_1011, results[1])
        expect_equals(0b0001_0111_0110, results[2])
        expect_equals(0b0000_1001_1011, results[3])
        expect_equals(0b0001_1011_1011, results[4])
        expect_equals(0b0001_0100_1011, results[5])
        expect_equals(0b0001_0101_0111, results[6])
        expect_equals(0b0001_0101_1011, results[7])
        expect_equals(0b0000_1111_0011, results[8])


def test_derive_allcandidates():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        candidates = testlib.getcollectionobj()
        res = testlib.callmacro(None, "DeriveCandidates", problem, candidates)
        expect(res[3] == 0)
