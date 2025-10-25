from sys import exec_prefix
from vbaunit_lib.testlib import gettestlib, expect, expect_equals, ignore, description  # noqa


def get_problemboard_beginner() -> list[list[int]]:
    problem = [
        [1, 2, 0, 4, 5, 0, 7, 0, 9],
        [4, 0, 6, 0, 8, 9, 1, 2, 0],
        [0, 8, 9, 1, 0, 3, 0, 5, 6],
        [2, 3, 0, 5, 6, 0, 8, 9, 1],
        [5, 6, 7, 8, 0, 1, 0, 3, 4],
        [8, 0, 1, 0, 3, 4, 5, 0, 7],
        [3, 4, 0, 6, 7, 8, 0, 1, 2],
        [6, 7, 8, 0, 1, 2, 3, 0, 5],
        [9, 1, 2, 3, 4, 5, 6, 7, 8],
    ]
    return problem


def get_problemboard_beginner_solved() -> list[list[int]]:
    solution = [
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
    return solution


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


def get_problemboard_expert_solved() -> list[list[int]]:
    # Corresponding solution
    solution = [
        [5, 8, 1, 2, 9, 4, 7, 6, 3],
        [9, 4, 3, 6, 7, 1, 8, 2, 5],
        [7, 2, 6, 3, 8, 5, 1, 9, 4],
        [3, 9, 8, 4, 1, 2, 5, 7, 6],
        [6, 5, 4, 7, 9, 3, 2, 1, 8],
        [2, 1, 7, 8, 5, 6, 9, 3, 4],
        [8, 7, 9, 5, 3, 2, 6, 4, 1],
        [1, 3, 5, 9, 4, 7, 3, 8, 2],
        [4, 6, 2, 1, 8, 9, 3, 5, 7],
    ]
    return solution


def test_fullprocess_beginner():
    testlib = gettestlib()
    problem = get_problemboard_beginner()
    solution = get_problemboard_beginner_solved()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm") as book:
        book.sheets("puzzle")["A1"].value = problem
        res = testlib.callmacro(None, "SolvePuzzle")
        expect_equals(0, res[1])
        expect_equals(solution, book.sheets("solution")["A1:I9"].value)
        expect_equals("Solved!", book.sheets("solution")["K1"].value)


def test_fullprocess_expert():
    testlib = gettestlib()
    problem = get_problemboard_expert()
    solution = get_problemboard_expert_solved()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm") as book:
        book.sheets("puzzle")["A1"].value = problem
        res = testlib.callmacro(None, "SolvePuzzle")
        expect_equals(0, res[1])
        expect_equals(solution, book.sheets("solution")["A1:I9"].value)
        expect_equals("Solved!", book.sheets("solution")["K1"].value)
