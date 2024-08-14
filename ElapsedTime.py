# -------------------------------------------------------------------------------
# Name:        ElapsedTime
# Purpose:     Measure the elapsed time between two events and print it in a human-readable format
#
#
# Author:      Gesuri
#
# Created:     August 14, 2024
# Updated:     August 14, 2024
#
# Copyright:   (c) Gesuri 2024
#
# -------------------------------------------------------------------------------

import datetime


def td_format(td_object):
    seconds = int(td_object.total_seconds())
    periods = [
        ('year',        60*60*24*365),
        ('month',       60*60*24*30),
        ('day',         60*60*24),
        ('hour',        60*60),
        ('minute',      60),
        ('second',      1),
    ]
    strings = []
    for period_name, period_seconds in periods:
        if seconds > period_seconds:
            period_value , seconds = divmod(seconds, period_seconds)
            has_s = 's' if period_value >= 1 else ''
            strings.append("%s %s%s" % (period_value, period_name, has_s))
    return ", ".join(strings)


class ElapsedTime:
    startTime = None
    endTime = None
    returnStr = True

    def __init__(self, returnStr=True):
        self.returnStr = returnStr
        self.start()

    @staticmethod
    def _current_():
        return datetime.datetime.now()

    def elapsed(self):
        if not self.endTime:
            self.endTime = self._current_()
        passedTime = self.endTime - self.startTime
        if self.returnStr:
            return td_format(passedTime)
        else:
            return passedTime

    def start(self):
        self.startTime = self._current_()
        self.endTime = None

    def end(self):
        self.endTime = self._current_()
        return self.elapsed()