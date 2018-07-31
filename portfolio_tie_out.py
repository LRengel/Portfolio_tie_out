import pandas as pd
import numpy as np
import argparse
import os
import time
# default values if none are given
ACH_DEFAULTCOLS = ('ContractNumber', 'CustomerName', 'Type', 'Bank Code', 'Amount', 'Program')
ACH_DEFAULTEXCEL = 'A,B,C,D,G,H'
ACH_DTYPES = {'Bank Code': np.object}
PORTFOLIO_BUYOUTCOLUMNNAMES = ('ContractNumber', 'CustomerName', 'Amount', 'Notes')
PORTFOLIOCOMULMNS = ('ContractNumber', 'CustomerName', 'Amount')
PORTFOLIOEXCEL = 'A,C,D'
DEFAULT_BANK_CODE = '999.99'


def add_extension(contract, extension):
    if type(contract) == float:
        return ' '
    elif len(contract) == 11:
        return '-'.join([extension, contract])
    else:
        return contract


def clean_ach(file, cols=ACH_DEFAULTCOLS, excelcolumns=ACH_DEFAULTEXCEL, bank_code=DEFAULT_BANK_CODE):
    """ Cleans the ACH spreadsheet and reformats it to a DataFrame for analysis

    Parameters

    ----------
    file : str
        The ach file you want to clean (must have a .xls or .xlsx extension)
    cols : tuple, optional
        An flag used to store to final DataFrame columns (default is ACH_DEFAULTCOLS a tuple of
        column names)
    excelcolumns : str, optional
        A flag used to specify which excel columns to load in from the ach file
        (default is ACH_DEFAULTEXCEL which is a str of excel column labels)
    bank_code : str, optional
        A flag used to filter all the ACH payments to compare

    Returns

    -------
    DataFrame
        A DataFrame that is has been cleaned to only include the porfolio payments and proper
        columns
    """
    df = pd.read_excel(file, usecols=excelcolumns, names=cols, dtype=ACH_DTYPES)
    df = df.loc[(df['Type'] == 'U') & (df['Bank Code'] == bank_code), :]
    df.reset_index(inplace=True)
    df.drop(columns=['index', 'Bank Code'], inplace=True)
    return df


def clean_portfolio(file, cut=9, rows=0, buyoutcolumns=PORTFOLIO_BUYOUTCOLUMNNAMES):
    """Cleans a portfolio file that has buyouts and reformats it to a DataFrame for analysis

    Parameters

    ----------
    file : str
        portfolio file you want to clean (must have a .xls or .xlsx extension)
    cut : int, optional
        a flag to split up the payments from the buyout section
    rows : int, optional
        rows from the top of the file to skip when it is read in
    buyoutcolumns: tuple, optional
        column names given to the cleaned DataFrame


    Returns

    -------
    DataFrame
        A DataFrame that has been cleaned to only include important information
    """
    df = pd.read_excel(file, rows)
    buyouts = df.loc[df.ContractNumber == 'Buyout', :].copy()
    df.dropna(axis=0, thresh=1)
    df = df.iloc[:cut, :]
    df.dropna(axis=1, inplace=True)
    buyouts = buyouts[['CustomerName', 'Amount', 'Unnamed: 3', 'ContractNumber']]
    buyouts.columns = buyoutcolumns
    buyouts.Amount = np.round(buyouts.Amount.astype(float), 2)
    newdf = df.merge(buyouts, on='ContractNumber', how='left')
    newdf.loc[newdf.Notes == 'Buyout', 'Amount_x'] = newdf.Amount_y
    newdf.drop(columns=['CustomerName_y', 'Amount_y'], inplace=True)
    newdf.fillna('', inplace=True)
    newdf.rename(columns={'Amount_x': 'Amount', 'CustomerName_x': 'CustomerName'}, inplace=True)
    return newdf


def clean_portfolio2(file, rows, footer):
    """Cleans the portfolio excel file with cancelled or replaced values and reformats it to a DataFrame

    Parameters

    ----------
    file : str
        portfolio file you want to clean (must have a .xls or .xlsx extension)
    rows : int
        rows from the top of the file to skip when it is read in
    footer : int
        rows from the bottom of the file to skip when the file is read in


    Returns

    -------
    DataFrame
        The reformated Dataframe ready to be compared with ach file read in
    """
    df = pd.read_excel(file, usecols=PORTFOLIOEXCEL, skip_footer=footer, skiprows=rows, names=PORTFOLIOCOMULMNS)
    df.ContractNumber.iloc[:9] = df.ContractNumber.apply(add_extension, args=('001',))
    df.ContractNumber.iloc[9:] = df.ContractNumber.apply(add_extension, args=('040',))
    df.dropna(axis=0, inplace=True)
    df = df.loc[df.CustomerName != 'Grand Total']
    df['Notes'] = df.Amount.where(cond=(df['Amount'] == 'Cancelled') | (df['Amount'] == 'Replaced'), other=0)
    df.loc[(df.Amount == 'Cancelled') | (df.Amount == 'Replaced'), 'Amount'] = 0
    df.Amount = df.Amount.astype(float)
    return df


def make_final_df(ach, port_df, portfolio_name):
    """ Merges the ACH DataFrame and Portfolio DataFrame together and compares the payments made


    Parameters

    ----------
    ach : DataFrame
        The ACH dataFrame created from the clean_ach function
    port_df : DataFrame
        The portfolio DataFrame created from one of the two clean_portfolio functions
    porfolio_name : str
        the name of the portfolio you are comparing payments with

    Returns

    -------
    DataFrame
        The merged DataFrame that show the differences between the files
    """
    new_col_names = {'CustomerName_x': 'CustomerName', 'Amount_x': 'Amount', 'Amount_y': portfolio_name}
    final_df = ach.merge(port_df, on='ContractNumber', how='inner')
    final_df.drop(columns=['Program', 'CustomerName_y'], inplace=True)
    final_df.rename(columns=new_col_names, inplace=True)
    final_df.insert(loc=5, column='Difference', value=final_df['Amount'] - final_df[portfolio_name])
    final_df = final_df.append(final_df.sum(numeric_only=True), ignore_index=True)
    final_df = final_df.rename(index={19: 'Totals'})
    final_df.fillna('', inplace=True)
    return final_df


parser = argparse.ArgumentParser()

parser.add_argument('ach_file', nargs='?', help='Enter a valid ach file excel file with extension .xls or .xlsx: ', type=str, default='ach_test.xlsx')
parser.add_argument('portfolio_file', nargs='?', help='Enter a valid portfolio file with extension .xls or .xlsx: ', type=str, default='portfolio_2.xlsx')
parser.add_argument('-dest', '-destination', help='Enter the file path to save the combine file:', type=str, default=os.getcwd())
parser.add_argument('-d', '--date', help="Enter the date for the file in 'MM-DD-YY' format:", type=str)
parser.add_argument('-b', '--buyouts', help='Run the program to deal with buyouts', action='store_true')
parser.add_argument('-prow', '--portfoliorows', nargs='?', help='Enter the rows from the top of the portfolio file you want to skip:', type=int, default=28)
parser.add_argument('-pfooter', '--portfoliofooterbuyout', nargs='?', help='Enter the rows from the bottom of the portfolio file to skip:', type=int, default=2)
parser.add_argument('-pname', '--portfolioname', nargs='?', help='Enter the name of the portfolio:', type=str, default='Clark LLC')


def main():
    args = parser.parse_args()
    ach_df = clean_ach(args.ach_file)
    if args.buyouts:
        portfolio_df = clean_portfolio(args.portfolio_file)
    else:
        portfolio_df = clean_portfolio2(args.portfolio_file, args.portfoliofooterbuyout, args.portfoliorows)
    print('Merging files...\n')
    final_df = make_final_df(ach_df, portfolio_df, args.portfolioname)
    time.sleep(1)
    print('file being saved here {}.xlsx'.format('_'.join([os.getcwd(), 'Portfolio_tie_out', args.date])))
    # final_df.to_excel('portfolio_tie_out{}.xlsx'.format(args.date))
    print('Below is you final df: \n')
    print(final_df)


if __name__ == '__main__':
    main()
