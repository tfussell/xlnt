import sys

def format(args):
    print()
    for line in sys.stdin.readlines():
        print('"' + line.rstrip().replace('"', '\\"') + '"')

if __name__ == '__main__':
    format(sys.argv)
