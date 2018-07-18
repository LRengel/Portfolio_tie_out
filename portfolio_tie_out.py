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


def add_extension(contract, extension):
    if type(contract) == float:
        return ' '
    elif len(contract) == 11:
        return '-'.join([extension, contract])
    else:
        return contract


def clean_ach(file, cols=ACH_DEFAULTCOLS, excelcolumns=ACH_DEFAULTEXCEL):
    ''' Cleans the ACH spreadsheet to grab the values needed for anlysis '''
    df = pd.read_excel(file, usecols=excelcolumns, names=cols, dtype=ACH_DTYPES)
    df = df.loc[(df['Type'] == 'U') & (df['Bank Code'] == '999.99'), :]
    df.reset_index(inplace=True)
    df.drop(columns=['index', 'Bank Code'], inplace=True)
    return df


def clean_portfolio(file, cut=9, rows=0, buyoutcolumns=PORTFOLIO_BUYOUTCOLUMNNAMES):
    ''' Cleans a the porfolio that includes buyouts'''
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


def clean_portfolio2(file, rows=2, footer=28):
    ''' Cleans a the porfolio that includes cancelled or replaced contracts'''
    df = pd.read_excel(file, usecols=PORTFOLIOEXCEL, skip_footer=footer, skiprows=rows, names=PORTFOLIOCOMULMNS)
    df.ContractNumber.iloc[:9] = df.ContractNumber.apply(add_extension, args=('001',))
    df.ContractNumber.iloc[9:] = df.ContractNumber.apply(add_extension, args=('040',))
    df.dropna(axis=0, inplace=True)
    df = df.loc[df.CustomerName != 'Grand Total']
    df['Notes'] = df.Amount.where(cond=(df['Amount'] == 'Cancelled') | (df['Amount'] == 'Replaced'), other=0)
    df.loc[(df.Amount == 'Cancelled') | (df.Amount == 'Replaced'), 'Amount'] = 0
    df.Amount = df.Amount.astype(float)
    return df


def make_final_df(ach, port_df, portfolio_name='Clark LLC'):
    ''' Combines both data frames into one'''
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
parser.add_argument('portfolio_file', nargs='?', help='Enter a valid portfolio file with extension .xls or .xlsx: ', type=str, default='portfolio2.xlsx')
parser.add_argument('-dest', '-destination', help='Enter the file path to save the combine file', type=str, default=os.getcwd())
parser.add_argument('-d', '--date', help="Enter the date for the file in 'MM-DD-YY' format", type=str)
parser.add_argument('-b', '--buyouts', help='Run the program to deal with buyouts', action='store_true')


def main():
    args = parser.parse_args()
    ach_df = clean_ach(args.ach_file)
    if args.buyouts:
        portfolio_df = clean_portfolio(args.portfolio_file)
    else:
        portfolio_df = clean_portfolio2(args.portfolio_file)
    print('Merging files...\n')
    final_df = make_final_df(ach_df, portfolio_df)
    time.sleep(1)
    print('file being saved here {}.xlsx'.format('_'.join([os.getcwd(), 'Portfolio_tie_out', args.date])))
    # final_df.to_excel('portfolio_tie_out{}.xlsx'.format(args.date))
    print('Below is you final df: \n')
    print(final_df)


if __name__ == '__main__':
    main()
