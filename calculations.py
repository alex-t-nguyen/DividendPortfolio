import datetime
from dateutil.relativedelta import *


def get_cost_basis(num_shares, buy_price):
    if buy_price is None or buy_price == '-':
        return float("NaN")
    else:
        return float(num_shares * buy_price)


def get_market_value(num_shares, market_price):
    if num_shares is None or num_shares == '-':
        return float("NaN")
    else:
        return float(num_shares * market_price)


def get_gain_loss(market_value, cost_basis):
    return market_value - cost_basis


def get_growth(gain_loss, cost_basis):
    return gain_loss / cost_basis


def get_yield_on_cost(buy_price, dividend_rate):
    # Return as decimal (excel changes to percentage)
    if buy_price is None or buy_price == '-':
        return float("NaN")
    else:
        return dividend_rate / buy_price


def get_annual_income(num_shares, dividend_rate):
    if num_shares is None or num_shares == '-':
        return float("NaN")
    else:
        return num_shares * dividend_rate


def get_last_qual_purchase_date(ex_date):
    try:
        date_object = datetime.datetime.strptime(ex_date, "%Y-%m-%d")
        hold_time = relativedelta(months = +3)
        last_buy_date = date_object + hold_time
        last_buy_date = last_buy_date.strftime("%m/%d/%Y")
        return last_buy_date
    except ValueError as e:
        return 'No data available'