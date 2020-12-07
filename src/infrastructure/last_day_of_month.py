from datetime import timedelta


def last_day_of_month(month):
    next_month = month.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)
