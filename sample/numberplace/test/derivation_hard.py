from vbaunit_lib.testlib import gettestlib, expect, expect_equals  # noqa


def get_problemboard_expert() -> list[list[int]]:
    # Very hard Sudoku puzzle (0 = blank)
    problem = [
        [0, 0, 0, 2, 0, 0, 0, 0, 3],
        [0, 4, 0, 0, 0, 0, 0, 0, 0],
        [7, 0, 0, 0, 0, 5, 0, 0, 0],
        [0, 0, 8, 0, 0, 0, 0, 7, 0],
        [0, 0, 0, 7, 0, 3, 0, 0, 0],
        [0, 1, 0, 0, 0, 0, 9, 0, 0],
        [0, 0, 0, 5, 0, 0, 0, 0, 1],
        [0, 0, 0, 0, 0, 0, 0, 8, 0],
        [4, 0, 0, 0, 0, 9, 0, 0, 0],
    ]
    return problem


def get_problemboard_expert_halfway1() -> list[list[int]]:
    # When the solver explores cell (2,8) and finds no candidates, it musut backtrack.
    problem = [
        [1, 5, 6, 2, 4, 7, 8, 9, 3],
        [2, 4, 3, 1, 6, 8, 5, 0, 0],
        [7, 0, 0, 0, 0, 5, 0, 0, 0],
        [0, 0, 8, 0, 0, 0, 0, 7, 0],
        [0, 0, 0, 7, 0, 3, 0, 0, 0],
        [0, 1, 0, 0, 0, 0, 9, 0, 0],
        [0, 0, 0, 5, 0, 0, 0, 0, 1],
        [0, 0, 0, 0, 0, 0, 0, 8, 0],
        [4, 0, 0, 0, 0, 9, 0, 0, 0],
    ]
    return problem


def test_derive_blockcandidates_1_1():
    testlib = gettestlib()
    problem = get_problemboard_expert()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        res = testlib.callmacro(None, "DeriveBlockCandidates", problem, 0, 0, 0)
        expect_equals(0, res[5])
        assert isinstance(res[4], int)
        result: int = res[4]  # blockcandidates argument ByRef
        expect_equals(0b000000000001, result & 0b0000_0000_0001)  # 1 allowed
        expect_equals(0b000000000010, result & 0b0000_0000_0010)  # 2 unabled
        expect_equals(0b000000000100, result & 0b0000_0000_0100)  # 3 unabled
        expect_equals(0, result & 0b0000_0000_1000)  # 4 unabled
        expect_equals(0b000000010000, result & 0b0000_0001_0000)  # 5 allowed
        expect_equals(0b000000100000, result & 0b0000_0010_0000)  # 6 allowed
        expect_equals(0, result & 0b0000_0100_0000)  # 7 unabled
        expect_equals(0b000010000000, result & 0b0000_1000_0000)  # 8 allowed
        expect_equals(0b000100000000, result & 0b0001_0000_0000)  # 9 allowed
