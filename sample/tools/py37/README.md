# pyexe.exe

[![Build status](https://ci.appveyor.com/api/projects/status/n18f0997k18x87lw/branch/master?svg=true)](https://ci.appveyor.com/project/manthey/pyexe/branch/master)

Here is a stand-alone version of python that is a single Windows executable.

It consists of the most recent versions of Python (with builds for 2.7, 3.5,
and 3.6 each in 32-bit and 64-bit versions), pywin32, psutil, six, pip, 
setuptools, and includes all packages that can be included without additional 
dlls, excepting tkinter.

See the appveyor script for build instructions.

## Installing other modules

Python is most useful with additional modules.  The stand-alone executable can use pip to install modules from pypi to the local directory.  For instance:

```bash
py36-64.exe -m pip install --no-cache-dir --target . --upgrade sympy
```

Use `-m pip` to run the pip module.  Use `--no-cache-dir` to avoid writing files to the user's data directory.  Use `--target .` to install to the current directory, allowing you to import the modules easily.  Use `--upgrade` to replace existing files, such as the common `bin` directory.  Note that using `--upgrade` will overwrite or discard existing files, which may not be what you want (the `bin` directory will end up with just files for the most recently installed package).

## Differences from installed Python

Although the stand-alone Python attempts to have the same features as a normally installed Python, there are some differences.

- If command line options are specified, there may be some differences in `sys.flags`, since it is read-only and cannot be altered after start.
- `PYTHONHOME` is ignored.  This option doesn't make sense for a stand-alone version.
- `-V` and `PYTHONVERBOSE` don't print exactly the same information as installed Python, partly because the verbosity is increased after some modules are already imported.
- `--check-hash-based-pycs` is ignored.  This option cannot be changed after the Python executable starts.
- `-R` and `PYTHONHASHSEED` are ignored.  These options cannot be changed after the Python executable starts.
- `PYTHONCASEOK` is not honored on Python 2.7.  It behaves as installed Python for Python 3.x, i.e., `-E` does not ignore it, but `-I` does, see [Python issue 16826](https://bugs.python.org/issue16826) for some discussion.
- Not all environment variables are handled, such as: `PYTHONIOENCODING`, `PYTHONFAULTHANDLER`, `PYTHONLEGACYWINDOWSFSENCODING`, `PYTHONLEGACYWINDOWSSTDIO`, `PYTHONMALLOC`, `PYTHONCOERCECLOCALE`, `PYTHONDEVMODE`.  Some of these are ignored; some are used and cannot be suppressed with `-E` or `-I`.  Many of these could be handled properly with additional work.
