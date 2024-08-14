# -------------------------------------------------------------------------------
# Name:        Log
# Purpose:     Log the a line to a file. The file is stored in path log
#
#
# Author:      Gesuri
#
# Created:     February 27, 2018
# Updated:     August 13, 2024
#
# Copyright:   (c) Gesuri 2018
#
# -------------------------------------------------------------------------------

# This Python code defines a class called "Log" that is used for logging messages into a file. Here's a breakdown of
#   the code:
#   Overall, this code provides a flexible logging functionality for writing messages to a file with timestamp and
#       optional printing to the standard output.
#
#   1. The code imports various modules from the Python standard library, including `os.path`, `os`, `sys`, `time`,
#       `datetime`, and `pathlib`.
#
#   2. The code defines a constant variable `PATH_LOGS` which represents the current working directory.
#
#   3. The code defines a function `getStrTime` that returns the current time as a string in a specified format.
#       The format is set to `TIMESTAMP_FORMAT` if provided, otherwise it uses the default format `%Y%m%d_%H%M%S`.
#       The function takes two optional arguments `utc` and `dst` which are used to indicate whether the time should
#       be returned in UTC or with daylight saving time.
#
#   4. The code defines a class called `Log` which is used for logging messages. The class has the following methods:
#       - `__init__(self, name=None, path=PATH_LOGS, timestamp=True, fprint=True)`: The constructor method initializes
#           the `Log` object. It takes optional arguments `name`, `path`, `timestamp`, and `fprint`. If `name` is not
#           provided, it uses the base name of the script file (`argv[0]`) as the default name. If the name is an empty
#           string, it sets the name to `'test.log'`. The `path` argument specifies the directory where the log file
#           should be created (default is `PATH_LOGS`). The `timestamp` argument is a boolean that indicates whether to
#           add a timestamp to each log entry (default is `True`). The `fprint` argument is a boolean that indicates
#           whether to print the log entries to the standard output (default is `True`).
#       - `setName(self, name)`: Sets the name of the log file.
#       - `setPath(self, path)`: Sets the directory path for the log file.
#       - `setTimeStamp(self, timestamp)`: Sets the timestamp flag.
#       - `setFprint(self, fprint)`: Sets the flag for printing log entries to standard output.
#       - `getName(self)`: Returns the name of the log file.
#       - `getPath(self)`: Returns the directory path for the log file.
#       - `getTimeStamp(self)`: Returns the timestamp flag.
#       - `getFprint(self)`: Returns the flag for printing log entries to standard output.
#       - `getFullPath(self)`: Returns the full path of the log file.
#       - `w(self, line, ow=False)`: Writes a log entry to the file. If `ow` is `True`, it overwrites the existing file;
#           otherwise, it appends to the file. The log entry can be a string or any other object that can be converted
#           to a string. If the log file doesn't exist, it creates a new file. If there is an IO error, it tries to
#           change the ownership of the file (using `sudo chown`) and retries the write operation. If there is still an
#           error, it logs an error message to a separate error log file.
#       - `ow(self, line)`: Overwrites the existing log file with a new log entry.
#       - `error(self, line)`: Writes an error log entry.
#       - `warn(self, line)`: Writes a warning log entry.
#       - `info(self, line)`: Writes an information log entry.
#       - `live(self, line)`: Writes a live log entry.
#       - `debug(self, line)`: Writes a debug log entry.
#       - `fatal(self, line)`: Writes a fatal log entry.
#       - `line(self, line)`: Writes a live log entry.
#
# The code also includes some helper methods (`_checkPath_` and `_checkName_`) that are used to validate the path and
#   name of the log file and ensure they meet the requirements.
#


# from os.path import splitext, basename
from os import system, getcwd
# from sys import argv
from time import localtime
from datetime import datetime, timedelta
from pathlib import Path
from colorama import just_fix_windows_console

just_fix_windows_console()

TIMESTAMP_FORMAT = '%Y%m%d_%H%M%S'


def getStrTime(formato=None, utc=False, dst=False):
    """Return the current time in a string with format conts.TIMESTAMP_FORMAT."""
    if formato is None:
        formato = TIMESTAMP_FORMAT
    if utc:
        return str(datetime.utcnow().strftime(formato))
    elif dst:
        return str(datetime.now().strftime(formato))
    else:
        if localtime().tm_isdst:
            return str((datetime.now() - timedelta(hours=1)).strftime(formato))
        else:
            return str(datetime.now().strftime(formato))


def pRed(skk):
    print(f"\033[91m{skk}\033[00m")


def pGreen(skk):
    print(f"\033[92m{skk}\033[00m")


def pYellow(skk):
    print(f"\033[93m {skk}\033[00m")


def pLightPurple(skk):
    print(f"\033[94m {skk}\033[00m")


def pPurple(skk):
    print(f"\033[95m {skk}\033[00m")


def pCyan(skk):
    print(f"\033[96m {skk}\033[00m")


def pLightGray(skk):
    print(f"\033[97m {skk}\033[00m")


def pBlack(skk):
    print(f"\033[98m {skk}\033[00m")


class Log:
    """Log the line into the file.  V20230822
          line:      line to print
          path:      path without file name (const.PATH_LOGS)
          timestamp: boolean to indicate if add timestamp (True)
          fprint:    boolean, print in file, print in stdio (True)
          sprint:    boolean, print in stdio (True)
    """
    path = None

    # name = None

    def __init__(self, path=None, timestamp=True, fprint=True, sprint=True):
        self.sprint = sprint
        if path is None:
            path = getcwd()
        self.path = Path(path)
        if sprint:
            self._checkPath_()
        self.timestamp = timestamp
        self.fprint = fprint

    def _checkPath_(self):
        if not self.path.exists():
            if self.path.suffix == '':  # if the path is a directory by checking if there is an extension
                self.path.mkdir(parents=True)
                self.path.joinpath('log.log')
            else:  # the path is a file but maybe the directory does not exist so create it
                self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.is_dir():
            self.path = self.path.joinpath(f'{self.path.name}.log')

    def setTimeStamp(self, timestamp):
        self.timestamp = timestamp

    def setFprint(self, fprint):
        self.fprint = fprint

    def setSprint(self, sprint):
        self.sprint = sprint
        if self.sprint:
            self._checkPath_()

    # def getName(self):
    #     return self.name

    def getPath(self):
        return self.path

    def getTimeStamp(self):
        return self.timestamp

    def getFprint(self):
        return self.fprint

    def getSpint(self):
        return self.sprint

    def getFullPath(self):
        return self.path

    def w(self, line, ow=False, color=None):
        """ write the line into the file """
        if ow:
            wo = 'w'
        else:
            wo = 'a'
        if type(line) != 'str':
            line = str(line)
        if len(line) > 0:
            if self.fprint:
                if not self.path.is_file():
                    f = self.path.open('w')
                else:
                    try:
                        f = self.path.open(wo)
                    except IOError:
                        system(f'sudo chown pi:pi {self.path}')
                        try:
                            f = self.path.open(wo)
                        except IOError:
                            system(f'sudo echo error writing {line} in file {self.path.name} >> errLog.log')
                            return
            now = getStrTime()
            if self.fprint or self.sprint:
                if self.timestamp:
                    msg = f'{now}, {line}'
                else:
                    msg = line
                if self.sprint:
                    if color:
                        color(msg)
                    else:
                        print(msg)
            if self.fprint:
                if line[-1] != '\n':
                    line += '\n'
                if self.timestamp:
                    f.write(f'{now},{line}')
                else:
                    f.write(line)
                f.flush()
                f.close()

    def ow(self, line):
        """ Overwrite the same file log """
        self.w(line, ow=True)

    def error(self, line):
        self.w(f'[Error]: {line}', color=pRed)

    def warn(self, line):
        self.w(f'[Warning]: {line}', color=pYellow)

    def info(self, line):
        self.w(f'[Info]: {line}', color=pGreen)

    def live(self, line):
        self.w(f'[Live]: {line}')

    def debug(self, line):
        self.w(f'[Debug]: {line}', color=pCyan)

    def fatal(self, line):
        self.w(f'[Fatal]: {line}', color=pPurple)
