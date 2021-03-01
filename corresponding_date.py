import datetime


def month_in_reference():
    """
    :return a Tuple of the corresponding Month and Year of SAFT
    Example: current month 1 (January) of 2021 returns 12 (December) of 2020
    """
    months = [n for n in range(1, 13)]
    current_date = datetime.date.today()
    current_month = current_date.timetuple()[1]
    last_month = months[current_month - 2]
    current_year = current_date.timetuple()[0]

    if last_month == 12:
        year = current_year - 1
    else:
        year = current_year

    return last_month, year
