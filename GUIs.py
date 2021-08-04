import xlwings as xw
import tkinter as tk
from DividendPortfolio import Column, get_start_row
from yahoofinancials import YahooFinancials

def insert_new_ticker():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    # Create GUI frame
    root = tk.Tk()
    root.geometry("300x100")
    root.title("Insert New Ticker")
    # Create GUI forms
    ticker_label = tk.Label(root, text='TICKER:')
    #ticker_label.pack(pady = 2, padx = 2)
    ticker_label.place(x=30, y=30)

    ticker_temp = tk.StringVar()
    ticker_symbol = ""
    ticker_entry = tk.Entry(root, text=ticker_temp, width = 30)
    ticker_entry.place(x=80, y=30)
    ticker_entry.focus()
    def submit_ticker():
        ticker_symbol = ticker_temp.get()
        root.destroy()
        index = find_alphabetical_position(sheet, str(ticker_symbol))
        if index == -1:
            return
        sheet.range(f"{index}:{index}").insert()
        sheet.range((f"{index}", Column.ticker.value)).value = ticker_symbol
    submit_button = tk.Button(root, text='Insert ticker', width=12, fg = 'black', command=submit_ticker)
    submit_button.place(x = 105, y = 60)
    
    # Display GUI
    root.mainloop()


def find_alphabetical_position(sht, ticker):
    ticker_start_row = get_start_row(sht)
    tickers = sht.range(ticker_start_row, Column.ticker.value).options(expand='down', numbers=str).value
    tickers.append(ticker)

    try:
        inserted_ticker_long_name = YahooFinancials(ticker).get_stock_quote_type_data()[ticker]["longName"]
    except KeyError as e:
        print(f"Ticker symbol [{ticker}] not found in YahooFinance.")
        return -1
    # If no tieckers, insert at fixed position (row 3)
    # If tickers in list, get all names from long_name column and insert alphabetically
    if not tickers:
        return 3    # Row 3 to skip header lines in file
    else:
        long_name_list = sht.range(ticker_start_row, Column.long_name.value).options(expand='down', numbers = str).value
        long_name_list.append(inserted_ticker_long_name)
        long_name_list = sorted(long_name_list, key=str.lower)
        index = long_name_list.index(inserted_ticker_long_name) + 3   # Add 3 to skip the head lines in file
        return index
    """
    for symb in tickers:
        long_name_list.append(YahooFinancials(symb).get_stock_quote_type_data()[symb]["longName"])
    sorted_tickers = sorted(long_name_list, key=str.lower)
    index = sorted_tickers.index(inserted_ticker_long_name) + 3 # Add 3 to skip the header lines in file
    print(sorted_tickers)
    return index
    """


if __name__ == "__main__":
    xw.Book("DividendPortfolio.xlsm").set_mock_caller()
    insert_new_ticker()

