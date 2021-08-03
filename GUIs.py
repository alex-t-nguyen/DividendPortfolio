import tkinter as tk

def insert_new_ticker():
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
    submit_button = tk.Button(root, text='Insert ticker', width=12, fg = 'black', command=submit_ticker)
    submit_button.place(x = 105, y = 60)
    
    # Display GUI
    root.mainloop()

