## Portfolio Analysis script

### Main Use
- To automate the cleaning and comparing of payments in an ACH file vs a portfolio file

### Inputs:

#### For ACH function:
- ACH file in string form
- Rows and footer values to skip as ints
- Optional column names for the dataframe as a tuple of strings
- Optional excel columns for the excel columns you want to pull in as a string

#### For Portfolio file:
- Portfolio in the form of a string
- List of column names as strings
- area to cut from the portfolio file that doesn't include buyouts
##### For Cancelled or Replace contracts:
	- The values will be turned to zero values and the value replaced or cancelled will be moved
	- To a notes column

### Outputs:
- One excel file showing the difference between the two spread sheets
- This excel file will be saved to a desired folder within the directory 


### Functions:
- clean_ach
- clean_portfolio
- clean_portfolio_2
- make_final_df

### Command line arguments:
- Default arguments:
	- ach_file
	- portfolio_file

- Optional arguments:
	- date
	- buyouts
	- destination

### Basic usage:
- From the Command line

- ` python portfolio_tie_out.py 'achfile.xlsx' 'portfoliofile.xlsx' -d 071819 `
