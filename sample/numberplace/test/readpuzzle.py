from xlwings import Sheet
from vbaunit_lib.testlib import gettestlib, expect  # noqa


def test_readboard():
    testlib = gettestlib()
    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm") as book:
        puzzlesheet: Sheet = book.sheets["puzzle"]
        problem = [
            [1, 2, 3, 4, 5, 6, 7, 8, 9],
            [4, 5, 6, 7, 8, 9, 1, 2, 3],
            [7, 8, 9, 1, 2, 3, 4, 5, 6],
            [2, 3, 4, 5, 6, 7, 8, 9, 1],
            [5, 6, 7, 8, 9, 1, 2, 3, 4],
            [8, 9, 1, 2, 3, 4, 5, 6, 7],
            [3, 4, 5, 6, 7, 8, 9, 1, 2],
            [6, 7, 8, 9, 1, 2, 3, 4, 5],
            [9, 1, 2, 3, 4, 5, 6, 7, 8],
        ]
        puzzlesheet["A1"].value = problem

        res = testlib.callmacro(None, "ReadBoard", None)

        board = (
            (1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0),
            (4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 1.0, 2.0, 3.0),
            (7.0, 8.0, 9.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0),
            (2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 1.0),
            (5.0, 6.0, 7.0, 8.0, 9.0, 1.0, 2.0, 3.0, 4.0),
            (8.0, 9.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0),
            (3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 1.0, 2.0),
            (6.0, 7.0, 8.0, 9.0, 1.0, 2.0, 3.0, 4.0, 5.0),
            (9.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0),
        )

        expect(res[1] == board)
