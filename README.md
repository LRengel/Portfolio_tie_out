# Portfolio Analysis Command Line Script

### Command line app for the cleaning and comparision of payments of two excel files


### Basic usage:
- From the Command line

- When no buyouts are included (default command line arguments are set to deal with no buyouts):

- ` python portfolio_tie_out.py  --date <str> `

- When buyouts are included:
- ` python portfolio_tie_out.py <achfile.xlsx> <portfoliofile.xlsx> --buyouts True --date <str> --cut <int>`


## Summary of Tech Stack:
This comand line app was built using **argparse** because I liked how easy it was to implement and it already comes in the standard library.

I chose **pandas** over **openpyxl** because pandas utilizes vectorization, which allows for faster code execution and avoids the loopoing that would've come from using openpxl.

## Functionality
This app takes two excel files cleans and combines them to create a new excel file that compares the payments between them. It also takes a previous month's final report to help with reformating the contract numbers of the portfolio file and the payment total sent from the company paying the portfolio. This script then gives you the option to create an excel file that has three sheets. One for the summary that includes the totals between the two excel files and notes on any cancelled payments or buyouts, another for the detailed comparison between the two excel files, and finally the cleaned portofile in order to tie out the two files 

## Testing
`pytest test_portfolio_tie_out.py `