import doctest
import unittest

import base
import templates.qiu as qiu


def load_tests(loader, tests, ignore):
    modules = [base, qiu]
    for module in modules:
        tests.addTests(doctest.DocTestSuite(module))
    return tests


if __name__ == "__main__":
    unittest.main()
