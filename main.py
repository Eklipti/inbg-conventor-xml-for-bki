import sys

import cli
import gui


def main():
    if len(sys.argv) == 1:
        gui.run()
    else:
        cli.run()


if __name__ == "__main__":
    main()
