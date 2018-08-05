import pandas as pd
import numpy as np
import argparse
import os
import time
# Default values if none are given
ACH_DEFAULTCOLS = ('ContractNumber', 'CustomerName', 'Type', 'Bank Code', 'Amount', 'Program')
ACH_DEFAULTEXCEL = 'A,B,C,D,G,H'
ACH_DTYPES = {'Bank Code': np.object}
PORTFOLIO_BUYOUTCOLUMNNAMES = ('ContractNumber', 'CustomerName', 'Amount', 'Notes')
PORTFOLIOCOMULMNS = ('ContractNumber', 'CustomerName', 'Amount')
PORTFOLIOEXCEL = 'A,C,D'
BANK_CODE = '999.99'


def clean_ach(file, cols=ACH_DEFAULTCOLS, excelcolumns=ACH_DEFAULTEXCEL, bank_code=BANK_CODE):
    """ Cleans the ACH spreadsheet and reformats it to a DataFrame for analysis

    Parameters

    ----------
    file : str
        The ach file you want to clean (must have a .xls or .xlsx extension)
    cols : tuple, optional
        An flag used to store to final DataFrame columns (default is ACH_DEFAULTCOLS)
    excelcolumns : str, optional
        A flag used to specify which excel columns to load in from the ach file
        (default is ACH_DEFAULTEXCEL)
    bank_code : str, optional
        A flag used to filter all the ACH payments to compare (default is BANK_CODE)

    Returns

    -------
    DataFrame
        A DataFrame that is has been cleaned to only include the portfolio payments and proper
        columns
    """
    df = pd.read_excel(file, usecols=excelcolumns, names=cols, dtype=ACH_DTYPES)
    df = df.loc[(df['Type'] == 'U') & (df['Bank Code'] == bank_code), :]
    df.reset_index(inplace=True)
    df.drop(columns=['index', 'Bank Code'], inplace=True)
    return df


def clean_portfolio(file, lastmonth, buyout=False, cut=0, rows=0, footer=0, buyoutcolumns=PORTFOLIO_BUYOUTCOLUMNNAMES):
    """Cleans portfolio files with or without buyouts and reformats it to a DataFrame

    Parameters

    ----------
    file : str
        portfolio file you want to clean (must have a .xls or .xlsx extension)
    lastmonth : str
        last month's portfolio file used reformat the contract number rows
        (must have a .xls or .xlsx extension)
    buyout : boolean
        specifies whether the file read in contains buyouts or not
    cut : int
        a flag to split up the payments from the buyout section
    rows : int
        rows from the top of the file to skip when the file is read in
    footer : int
        rows from the bottom of the file to skip when the file is read in

    Returns

    -------
    DataFrame
        The reformatted DataFrame ready to be compared with ach file read in
    """
    if buyout:
        df = pd.read_excel(file)
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
    else:
        df = pd.read_excel(file, usecols=PORTFOLIOEXCEL, skip_footer=footer, skiprows=rows, names=PORTFOLIOCOMULMNS)
        df.dropna(axis=0, inplace=True)
        df.ContractNumber = pd.read_excel(lastmonth, usecols='A').values
        df['Notes'] = df.Amount.where(cond=(df['Amount'] == 'Cancelled') | (df['Amount'] == 'Replaced'), other=0)
        df.loc[(df.Amount == 'Cancelled') | (df.Amount == 'Replaced'), 'Amount'] = 0
        df.Amount = df.Amount.astype(float)
        return df


def make_final_df(ach, port_df, portfolio_name):
    """ Merges the ACH and Portfolio DataFrame together and compares the payments made

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
        The merged DataFrame that shows the differences between the files
    """
    new_col_names = {'CustomerName_x': 'CustomerName', 'Amount_x': 'Amount', 'Amount_y': portfolio_name}
    final_df = ach.merge(port_df, on='ContractNumber', how='inner')
    final_df.drop(columns=['Program', 'CustomerName_y'], inplace=True)
    final_df.rename(columns=new_col_names, inplace=True)
    final_df.insert(loc=5, column='Difference', value=final_df['Amount'] - final_df[portfolio_name])
    final_df = final_df.append(final_df.sum(numeric_only=True), ignore_index=True)
    final_df = final_df.rename(index={final_df.shape[0] - 1: 'Totals'})
    final_df.fillna('', inplace=True)
    return final_df


parser = argparse.ArgumentParser()

parser.add_argument('ach_file', nargs='?', help='Enter a valid ach file excel file with extension .xls or .xlsx: ', type=str, default='ach_test.xlsx')
parser.add_argument('portfolio_file', nargs='?', help='Enter a valid portfolio file with extension .xls or .xlsx: ', type=str, default='portfolio_2.xlsx')
parser.add_argument('lastmonth_file', nargs='?', help='Enter a valid final excel file to be used to get the proper contract numbers', type=str, default='Portfolio_tie_out_060218.xlsx')
parser.add_argument('--cut', nargs='?', help='Enter the row to split the portfolio file between the buyouts and regular columns', type=int, default=0)
parser.add_argument('--destination', help='Enter the file path to save the combined file', type=str, default=os.getcwd())
parser.add_argument('--date', help="Enter the date for the file in 'MM-DD-YY' format", type=str)
parser.add_argument('--buyouts', help='Run the program to deal with buyouts', type=bool, default=False)
parser.add_argument('--portfoliorows', nargs='?', help='Enter the rows from the top of the portfolio file to skip:', type=int, default=2)
parser.add_argument('--pfooter', nargs='?', help='Enter the rows from the bottom of the portfolio file to skip', type=int, default=9)
parser.add_argument('--portfolioname', nargs='?', help='Enter the name of the portfolio', type=str, default='Clark LLC')


def main():
    args = parser.parse_args()
    ach_df = clean_ach(args.ach_file)
    portfolio_df = clean_portfolio(args.portfolio_file, args.lastmonth_file, args.buyouts, args.cut, args.portfoliorows, args.pfooter)
    print('Merging files...\n')
    final_df = make_final_df(ach_df, portfolio_df, args.portfolioname)
    time.sleep(1)
    print('file being saved here {}.xlsx'.format('_'.join([os.getcwd(), 'Portfolio_tie_out', args.date])))
    # final_df.to_excel('portfolio_tie_out{}.xlsx'.format(args.date))
    print('Below is you final df: \n')
    print(final_df)


if __name__ == '__main__':
    main()
