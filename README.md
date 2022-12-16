# Wall Street Yearly Stock Reader

This is a VBA module that attempts to read and summarize stock trading data from a given year. 

# Capabilities

This sheet reads daily stock data data from the year 2020 (and attempts to read data from other stock years.) It will output each stock's change in price along with its percent change and total volume traded. In addition, the largest percent increase, largest percent decrease, and largest total volume are also found. All values are presented from columns I to Q, so that range should be kept clear before the module is run.


# Limitations
The sheet read must follow a very specific format. Stocks must be sorted by date followed by ticker. In addition, opening price must be on column C, closing price on column F, and total trading volume on column G. The data must begin on row 2, with row 1 being reserved for the header. The date of cell (B2) must be in the format "20200102". If the data sheet is not in this format or from the year 2020, it will still attempt to output to output data but the information will not be correct. 

Considering the current limitations of the module, it should be considered a beta and for demonstration purposes only. 





