# VBA-challenge

On each sheet:
  pre-check: if all sheets have same column/header
  generated summary1 for each ticker
    final columns: ticker, yearly change, percent change, total stock volume
      -helper columns (open date, open price, close date, close price) deleted 
      -I understand it might be easier when looking for open/close price only on 1/2 and 12/31
      -if in read world, I would use database tool to solve the problem
  format all columns as requested
    
  generated summary2 based on summary1
    ticker name & value of greatest increase %
    ticker name & value of greatest decrease %
    ticker name & value of greatest total stock volume

When test this VBA script in the alphabetical_testing file, it finished within 2min.
When run against the big yearly stock data, it took much longer time.

Results are stored as screenshots png files named by years.
  
