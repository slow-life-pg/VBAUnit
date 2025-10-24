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


def get_problemboard_beginner_tuple() -> tuple[
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
    problem = (
        (1, 2, 0, 4, 5, 0, 7, 0, 9),
        (4, 0, 6, 0, 8, 9, 1, 2, 0),
        (0, 8, 9, 1, 0, 3, 0, 5, 6),
        (2, 3, 0, 5, 6, 0, 8, 9, 1),
        (5, 6, 7, 8, 0, 1, 0, 3, 4),
        (8, 0, 1, 0, 3, 4, 5, 0, 7),
        (3, 4, 0, 6, 7, 8, 0, 1, 2),
        (6, 7, 8, 0, 1, 2, 3, 0, 5),
        (9, 1, 2, 3, 4, 5, 6, 7, 8),
    )
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


def get_problemboard_expert():
    problem = [
        [8, 0, 0, 0, 1, 0, 0, 2, 0],
        [0, 0, 0, 8, 0, 0, 4, 0, 0],
        [0, 0, 7, 0, 0, 3, 0, 0, 8],
        [5, 0, 0, 0, 3, 0, 0, 0, 2],
        [0, 0, 0, 0, 0, 8, 0, 4, 0],
        [0, 4, 0, 5, 0, 0, 0, 0, 0],
        [0, 0, 0, 0, 8, 1, 0, 9, 0],
        [9, 0, 0, 0, 0, 0, 2, 0, 0],
        [0, 0, 4, 9, 0, 0, 0, 0, 5],
    ]
    return problem


def get_problemboard_expert_solved():
    solution = [
        [8, 3, 5, 4, 1, 6, 9, 2, 7],
        [2, 9, 6, 8, 5, 7, 4, 3, 1],
        [4, 1, 7, 2, 9, 3, 6, 5, 8],
        [5, 6, 9, 1, 3, 4, 7, 8, 2],
        [1, 2, 3, 6, 7, 8, 5, 4, 9],
        [7, 4, 8, 5, 2, 9, 1, 6, 3],
        [6, 5, 2, 7, 8, 1, 3, 9, 4],
        [9, 8, 1, 3, 4, 5, 2, 7, 6],
        [3, 7, 4, 9, 6, 2, 8, 1, 5],
    ]
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


def test_search_leaf_solved():
    testlib = gettestlib()
    problem = get_problemboard_beginner_solved()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)  # It’s already filled in, but that doesn’t matter.
        expect_equals(0, res[3])
        expect_equals(-1, res[0])  # SOLVED


def test_search_phi():
    testlib = gettestlib()
    problem = get_problemboard_impossible()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)
        expect_equals(0, res[3])
        expect_equals(0, res[0])  # PHI


def test_search_solve_onestep():
    testlib = gettestlib()
    problem = get_problemboard_beginner_solved()

    problem[8][8] = 0  # target

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)
        expect_equals(0, res[3])
        expect_equals(-1, res[0])  # SOLVED


def test_search_solve_twostep():
    testlib = gettestlib()
    problem = get_problemboard_beginner_solved()

    problem[4][3] = 0  # target1
    problem[8][8] = 0  # target2

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)
        expect_equals(0, res[3])
        expect_equals(-1, res[0])  # SOLVED


def test_search_solve_fullstep():
    testlib = gettestlib()
    problem = get_problemboard_beginner()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)
        expect_equals(0, res[3])
        expect_equals(-1, res[0])  # SOLVED


@description("There are too many branches, but the process finishes within a practical amount of time.")
def test_search_solve_difficult():
    testlib = gettestlib()
    problem = get_problemboard_expert()

    with testlib.runapp("..\\product\\NumberPlaceSolver.xlsm"):
        solution = testlib.getdynamicarray(0, 8, 0, 8)
        res = testlib.callmacro(None, "SearchSolution", problem, solution)
        expect_equals(0, res[3])
        expect_equals(-1, res[0])  # SOLVED
