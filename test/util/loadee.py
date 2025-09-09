from vbaunit_lib.testlib import description, ignore


def return_test_str() -> str:
    return "This is test function in loadee module."


@description("This is a test function that runs.")
def test_function_run():
    print("test_function_run executed.")


@ignore
def test_function_ignore():
    print("test_function_ignore executed.")
