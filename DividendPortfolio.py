import xlwings as xw
from yahoofinancials import YahooFinancials
import pandas as pd
from enum import Enum
import calculations as calc

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.name = "Portfolio"
    df = pull_stock_data(sheet)
    write_data_to_excel(df, sheet)


def pull_stock_data(sht):
    """
    Steps:
    1) Create an empty DataFrame
    2) Get tickers
    3) Iterate over tickers, pull data from Yahoo Finance & add data to dictionary "new row"
    4) Append "new row" to DataFrame
    5) Return DataFrame
    """
    # Create an empty DataFrame
    df = pd.DataFrame()
    
    # Get tickers
    tickers = get_tickers(sht)
    if tickers:
        for ticker in tickers:
            # Pull data from YahooFinance
            data = YahooFinancials(ticker)
            market_price = data.get_current_price()
            if market_price is None:
                print(f"Ticker: {ticker} not found on Yahoo Finance.")
                df = df.append(pd.Series(dtype=str), ignore_index=True)
            else:
                # get_stock_quote_type_data() returns 
                # dictionary in dictionary of ({"Ticker": {data}}) - ticker is key
                try:
                    long_name = data.get_stock_quote_type_data()[ticker]["longName"]
                except (TypeError, KeyError):
                    long_name = None
                # Calculations
                avg_cost_basis = calc.get_cost_basis(get_num_shares(sht, ticker), get_buy_price(sht, ticker))
                market_value = calc.get_market_value(get_num_shares(sht, ticker), market_price)
                gain_loss = calc.get_gain_loss(market_value, avg_cost_basis)
                growth = calc.get_growth(gain_loss, avg_cost_basis)

                # Create new row with data for stock
                new_row = {
                    "long_name": long_name,
                    "ticker": ticker,
                    "shares": get_num_shares(sht, ticker),
                    "additional_shares": "-",
                    "buy_price": get_buy_price(sht, ticker),
                    "new_buy_price": "-",
                    "market_price": market_price,
                    "avg_cost_basis": avg_cost_basis,
                    "market_value": market_value,
                    "gain_loss": gain_loss,
                    "growth": round(growth, 2),
                    "annual_dividend": float(data.get_dividend_rate()),
                    "dividend_yield": round(data.get_dividend_yield(), 2),
                    "yield_on_cost": round(calc.get_yield_on_cost(get_buy_price(sht, ticker), data.get_dividend_rate()), 2),
                    "annual_income": round(calc.get_annual_income(get_num_shares(sht, ticker), data.get_dividend_rate()), 2),
                    "ex_date": data.get_exdividend_date(),
                    "last_qual_purchase_date": calc.get_last_qual_purchase_date(data.get_exdividend_date()),
                }
                # Append data for stock to dataframe
                df = df.append(new_row, ignore_index=True)
    
    # Return data to excel
        return df
    return pd.DataFrame()


def write_data_to_excel(df, sht):
    manual_input_cols = ["ticker"]
    if not df.empty:
        options = dict(index=False, header=False)
        for data in Column:
            if not data.name in manual_input_cols:
                start_row = get_start_row(sht)
                    # Assigns the Pandas Series (entire column) to the excel sheet at sht.range(row, col)
                    # Doesn't iterate over the rows, simply assigns the top of the column to the starting point in excel sheet
                    # and the rest of the rows in the column get pasted below
                sht.range(start_row, data.value).options(**options).value = df[data.name]
        return None


def get_tickers(sht):
    ticker_start_row = get_start_row(sht)

    # Range is row (2), column (2) if using only numbers (Can also do just B2 for the entire column)
    tickers = sht.range(ticker_start_row, Column.ticker.value).options(expand='down', numbers=str).value
    return tickers


def get_start_row(sht):
    # "TICKER" name range is manually put in excel sheet first for program to find
    return sht.range("TICKER").row + 1 # Plus 1 row after the heading


def get_last_row(sht):
    # Go up until you hit a non-empty cell and get row of that cell
    return sht.range(sht.cells.last_cell.row, Column.ticker.value).end('up').row


def get_num_shares(sht, ticker):
    for row in range(2, get_last_row(sht) + 1):
        if sht.range((row, Column.ticker.value)).value == ticker:
            orig_shares = sht.range((row, Column.shares.value)).options(numbers=int).value
            new_shares = 0
            if sht.range((row, Column.additional_shares.value)).value != '-' and sht.range((row, Column.additional_shares.value)).value is not None:
                new_shares = sht.range((row, Column.additional_shares.value)).options(numbers=int).value
            return orig_shares + new_shares
    return float('NaN')


def get_buy_price(sht, ticker):
    for row in range(2, get_last_row(sht) + 1):
        if sht.range((row, Column.ticker.value)).value == ticker:
            orig_buy_price = float(sht.range((row, Column.buy_price.value)).options(numbers=float).value)
            new_buy_price = 0
            if sht.range((row, Column.new_buy_price.value)).options(numbers=float).value != '-' and sht.range((row, Column.new_buy_price.value)).value is not None:
                new_buy_price = float(sht.range((row, Column.new_buy_price.value)).options(numbers=float).value)
            if new_buy_price != 0:
                return (orig_buy_price + new_buy_price) / 2
            else:
                return orig_buy_price
    return float('NaN')

   
class Column(Enum):
    """ Column Name Translation from Excel, 1 = Column A, 2 = Colmumn B, ... """

    long_name = 1
    ticker = 2
    shares = 3
    additional_shares = 4
    buy_price = 5
    new_buy_price = 6
    market_price = 7
    avg_cost_basis = 8
    market_value = 9
    gain_loss = 10
    growth = 11
    annual_dividend = 12
    dividend_yield = 13
    yield_on_cost = 14
    annual_income = 15
    ex_date = 16
    last_qual_purchase_date = 17


if __name__ == "__main__":
    xw.Book("DividendPortfolio.xlsm").set_mock_caller()
    main()

