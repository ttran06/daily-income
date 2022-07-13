"""
Entrypoint module, in case you use `python -mdaily_income`.


Why does this file exist, and why __main__? For more info, read:

- https://www.python.org/dev/peps/pep-0338/
- https://docs.python.org/2/using/cmdline.html#cmdoption-m
- https://docs.python.org/3/using/cmdline.html#cmdoption-m
"""
import argparse
from pathlib import Path

from daily_income.cli import run


def parse_args():
    parser = argparse.ArgumentParser()

    parser.add_argument("month", type=int, help="month of income workbook")
    parser.add_argument("year", type=int, help="year of income workbook")
    parser.add_argument("--output", type=Path, default=".", help="Output path")

    return parser.parse_args()


def main():
    args = parse_args()

    run(args.month, args.year, args.output)


if __name__ == "__main__":
    main()
