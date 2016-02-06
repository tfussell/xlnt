"Benchmark some different implementations for cells"

from openpyxl.compat import range

from openpyxl.cell import Cell
from openpyxl.cell.read_only import ReadOnlyCell
from memory_profiler import memory_usage
import time


def standard():
    c = Cell(None, "A", "0", None)

def iterative():
    c = ReadOnlyCell(None, None, None, 'n')

def dictionary():
    c = {'ws':'None', 'col':'A', 'row':0, 'value':1}


if __name__ == '__main__':
    initial_use = memory_usage(proc=-1, interval=1)[0]
    for fn in (standard, iterative, dictionary):
        t = time.time()
        container = []
        for i in range(1000000):
            container.append(fn())
        print("{0} {1} MB, {2:.2f}s".format(
            fn.func_name,
            memory_usage(proc=-1, interval=1)[0] - initial_use,
            time.time() - t))
