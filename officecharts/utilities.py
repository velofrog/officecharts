import datetime as dt
from collections.abc import Iterable
import dateutil.parser
import pandas as pd
import numpy as np


def is_iterable(obj: Iterable[object] | object) -> bool:
    """Check if obj is iterable. Excludes string

    Args:
        obj (any): Any object

    Returns:
        bool: If obj is iterable
    """

    try:
        if isinstance(obj, str):
            return False

        iter(obj)
        return True

    except TypeError:
        return False


def is_datetime(obj: Iterable[object] | object) -> bool:
    """Checks if obj is, or can be coerced, to a datetime object

    Args:
        obj (any): Any object

    Returns:
        bool: If obj can be converted to a datetime object
    """

    if obj is None:
        return False

    if is_iterable(obj):
        return all([is_datetime(e) for e in obj])
    else:
        return as_datetime(obj) is not None


def as_datetime(obj: object) -> list[dt.datetime] | None:
    """Converts obj, if necessary, to a datetime object or list of datetime objects

    Args:
        obj (any): Any object or iterable

    Returns:
        any: datetime object(s)
    """

    if obj is None:
        return None

    if is_iterable(obj):
        return [as_datetime(e) for e in obj]

    try:
        if isinstance(obj, pd.Timestamp):
            return obj.to_pydatetime()

        if isinstance(obj, np.datetime64):
            return obj.astype(dt.datetime)

        if isinstance(obj, dt.date):
            return dt.datetime(obj.year, obj.month, obj.day)

        if isinstance(obj, dt.datetime):
            return obj

        if isinstance(obj, str):
            return dateutil.parser.parse(obj)

    except (TypeError, AttributeError, ValueError):
        return None

    return None


def excel_datetime(obj: list[dt.datetime] | dt.datetime) -> float | None:
    """Converts datetime obj to excel numeric date/time value

    Args:
        obj (datetime.datetime): A datetime object

    Returns:
        float: Excel numeric date/time value
    """

    if is_iterable(obj):
        return [excel_datetime(e) for e in obj]

    if not isinstance(obj, dt.datetime) and isinstance(obj, dt.date):
        obj = dt.datetime(obj.year, obj.month, obj.day)

    if not isinstance(obj, dt.datetime):
        return None

    value = (obj - dt.datetime(1900, 1, 1)) / dt.timedelta(days=1) + 1
    # Excel treats 1900 as a leap year.
    if obj >= dt.datetime(1900, 3, 1):
        value += 1

    return value
