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


def get_problemboard_beginner_solved_tuple() -> tuple[
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
    tuple[int, int, int, int, int, int, int, int, int],
]:
    solution = (
        (1, 2, 3, 4, 5, 6, 7, 8, 9),
        (4, 5, 6, 7, 8, 9, 1, 2, 3),
        (7, 8, 9, 1, 2, 3, 4, 5, 6),
        (2, 3, 4, 5, 6, 7, 8, 9, 1),
        (5, 6, 7, 8, 9, 1, 2, 3, 4),
        (8, 9, 1, 2, 3, 4, 5, 6, 7),
        (3, 4, 5, 6, 7, 8, 9, 1, 2),
        (6, 7, 8, 9, 1, 2, 3, 4, 5),
        (9, 1, 2, 3, 4, 5, 6, 7, 8),
    )
    return solution


def test_fullprocess_beginner():
    testlib = gettestlib()
    problem = get_problemboard_beginner()
    solution = get_problemboard_beginner_solved_tuple()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm") as book:
        book.sheets("puzzle")["A1"].value = problem
        res = testlib.callmacro(None, "SolvePuzzle")
        expect_equals(0, res[2])
        expect_equals(solution, book.sheets("solution")["A1:I9"].value)
        expect_equals("Solved!", book.sheets("solution")["K1"].value)
