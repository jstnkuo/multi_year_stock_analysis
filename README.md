# multi_year_stock_analysis
The following VBA script can be run to analyze stock data across all worksheets with in an Excel workbook

*IMPORTANT*
The limitation of this script:
can only be run if and only if 
  1) the stocks are sorted by stock name in the first column
  2) the stocks are sorted by increasing date in the second column
  3) opening price is in the third column
  4) closing price is in the sixth column
  5) stock volume is in the seventh column


The script will return the following of each stock in the data set:
  1) yearly change in column J
  2) percentage change in column K
  3) total stock volume in column L

Additionally, it calculates:
  1) the greatest % increase and value of the stock within the worksheet
  2) the greatest % decrease and value of the stock within the worksheet
  3) the value of the greatest total volume of the stock within the worksheet

Screenshots of examples of results are uploaded for reference.
