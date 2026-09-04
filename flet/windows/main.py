"""Flet packaging entry point for the Windows application.

``flet build linux .`` expects an application entry module in the project
directory.  The operational CLI remains in ``run.py``; importing and invoking
it here keeps the packaged and development launch paths identical.
"""

from run import main


if __name__ == "__main__":
    raise SystemExit(main())
