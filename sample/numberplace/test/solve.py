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


def get_problemboard_impossible() -> list[list[int]]:
    problem = [
        [1, 2, 0, 4, 5, 0, 7, 0, 9],
        [4, 0, 6, 0, 8, 9, 1, 2, 0],
        [0, 8, 9, 1, 0, 3, 0, 5, 6],
        [2, 3, 0, 5, 6, 0, 8, 9, 1],
        [5, 6, 7, 8, 2, 1, 0, 3, 4],
        [8, 0, 1, 0, 3, 4, 5, 0, 7],
        [3, 4, 0, 6, 7, 8, 0, 1, 2],
        [6, 7, 8, 0, 1, 2, 3, 0, 5],
        [9, 1, 2, 3, 4, 5, 6, 7, 8],
    ]
    return problem


def test_solve_beginner():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "Solve", problem, solution)
        expect_equals(0, res[3])
        expect_equals("Solved!", res[0])  # SOLVED
        expect_equals(get_problemboard_beginner_solved_tuple(), res[2])


@description("There are too many branches, but the process will finish within a practical amount of time.")
def test_solve_impossible():
    testlib = gettestlib()
    problem = get_problemboard_impossible()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "Solve", problem, solution)
        expect_equals(0, res[3])
        expect_equals("No solution.", res[0])  # UNSOLVED
