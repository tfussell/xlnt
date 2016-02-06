import os
import sys
import timeit

import openpyxl


def reader(optimised):
    """
    Loop through all cells of a workbook
    """
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "very_large.xlsx")
    wb = openpyxl.load_workbook(src, use_iterators=optimised)
    ws = wb.active
    rows = ws.iter_rows()
    for r, row in enumerate(rows):
        for c, col in enumerate(row):
            pass
    print((r+1)* (c+1), "cells")

def timer(fn):
    """
    Create a timeit call to a function and pass in keyword arguments.
    The function is called twice, once using the standard workbook, then with the optimised one.
    Time from the best of three is taken.
    """
    print("lxml", openpyxl.LXML)
    result = []
    for opt in (False, True,):
        print("Workbook is {0}".format(opt and "optimised" or "not optimised"))
        times = timeit.repeat("{0}({1})".format(fn.__name__, opt),
                              setup="from __main__ import {0}".format(fn.__name__),
                              number = 1,
                              repeat = 3
        )
        print("{0:.2f}s".format(min(times)))
        result.append(min(times))
    std, opt = result
    print("Optimised takes {0:.2%} time\n".format(opt/std))
    return std, opt


if __name__ == "__main__":
    timer(reader)
