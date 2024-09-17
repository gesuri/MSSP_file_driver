# -------------------------------------------------------------------------------
# Name:        ElapsedTime
# Purpose:     Measure the elapsed time between two events and print it in a human-readable format
#
# Author:      Gesuri
#
# Created:     August 14, 2024
# Updated:     August 14, 2024
#
# Copyright:   (c) Gesuri 2024
# -------------------------------------------------------------------------------

import datetime


def td_format(td_object):
    """
    Converts a timedelta object into a human-readable string.

    Args:
        td_object (datetime.timedelta): A timedelta object representing the elapsed time.

    Returns:
        str: A string representing the elapsed time in years, months, days, hours, minutes, and seconds.
    """
    # Convert the total seconds from the timedelta object into an integer.
    seconds = int(td_object.total_seconds())

    # Define the periods in seconds for years, months, days, hours, minutes, and seconds.
    periods = [
        ('year', 60 * 60 * 24 * 365),
        ('month', 60 * 60 * 24 * 30),
        ('day', 60 * 60 * 24),
        ('hour', 60 * 60),
        ('minute', 60),
        ('second', 1),
    ]

    strings = []
    # Loop through each period and calculate the number of each period in the total seconds.
    for period_name, period_seconds in periods:
        if seconds > period_seconds:
            # Calculate the number of periods and update the remaining seconds.
            period_value, seconds = divmod(seconds, period_seconds)
            has_s = 's' if period_value >= 1 else ''
            # Append the period and its count to the result string.
            strings.append("%s %s%s" % (period_value, period_name, has_s))

    # Join the periods into a single string and return it.
    return ", ".join(strings)


class ElapsedTime:
    """
    A class used to measure the elapsed time between two events.

    Attributes:
        startTime (datetime.datetime): The time when the timer started.
        endTime (datetime.datetime): The time when the timer ended.
        returnStr (bool): Flag to indicate whether to return the elapsed time as a string or a timedelta object.
    """

    startTime = None
    endTime = None
    returnStr = True

    def __init__(self, returnStr=True):
        """
        The constructor for the ElapsedTime class.

        Args:
            returnStr (bool, optional): Whether to return the elapsed time as a string. Defaults to True.
        """
        self.returnStr = returnStr
        self.start()

    @staticmethod
    def _current_():
        """
        Get the current datetime.

        Returns:
            datetime.datetime: The current datetime.
        """
        return datetime.datetime.now()

    def elapsed(self):
        """
        Calculate the elapsed time from start to end.

        Returns:
            str or datetime.timedelta: The elapsed time as a string or timedelta object based on the returnStr flag.
        """
        if not self.endTime:
            # If endTime is not set, set it to the current time.
            self.endTime = self._current_()

        # Calculate the time difference between startTime and endTime.
        passedTime = self.endTime - self.startTime

        # Return the elapsed time as a string or timedelta object.
        if self.returnStr:
            return td_format(passedTime)
        else:
            return passedTime

    def start(self):
        """
        Start or restart the timer by setting the startTime to the current time.
        """
        self.startTime = self._current_()
        self.endTime = None

    def end(self):
        """
        End the timer by setting the endTime to the current time.

        Returns:
            str or datetime.timedelta: The elapsed time as a string or timedelta object based on the returnStr flag.
        """
        self.endTime = self._current_()
        return self.elapsed()
