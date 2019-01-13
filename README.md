# Portfolio Analysis command line script

### Command line app for the cleaning and comparision of payments of two excel files


### Basic usage:
- From the Command line

- When no buyouts are included:

- ` python portfolio_tie_out.py 'achfile.xlsx' 'portfoliofile.xlsx' --date <date> `

- When buyouts are included:
- ` python portfolio_tie_out.py 'achfile.xlsx' 'portfoliofile.xlsx' --date <date> --buyout True`


## Summary of Tech Stack:
This comand line app was built using **argparse** because I liked how easy it was to implement and it already comes in the standard library.

I chose **pandas** over **openpyxl** because pandas utilizes vectorization, which allows for faster code execution and avoids the loopoing that would've come from using openpxl.

## Functionality
This app takes two excel files cleans and combines them to create a new excel file that compares the payments. This excel file contains three sheets one sheet for payment comparison between the excel files, another sheet for special notes pertaining to any cancelled or buyouts for each contract, and a sheet for the summary of the totals

## Testing
`pytest test_portfolio_tie_out.py `