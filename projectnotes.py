'''
Project Initials Notes:

	Dummy Data Criteria:

		ACH Excel File:
			Unclean:
				-Initial Columns:
					ContractNumber, CustomerName, Type, Bank Code, Other Amount, Some Amount, Amount,
						Program
				-columns types: objects, and float32 values, includes zero rows
				-Includes blank rows between ACH batches
			Clean:
				Columns: ['ContractNumber', 'CustomerName', 'Type', 'Amount']
				dtypes: object, object, categorical, float64
				Removed: Zero Value rows, N values, Bank Code

		Portfolio data:
			Unclean:
				Columns: ['ContractNumber', 'Company Contract', 'Customer Name', 'Date of payment']
				dtypes: object, object, object, float32

				Includes: -Amounts that have replaced or Cancelled as values
						  -Also includes notes at the bottom that need to be skipped
						  -Potentially buyout notes
						  -potentially contract numbers that are missing the leading 3 digits

		Both spreadsheets should have the same contract numbers, but some values may differ and include

	Functions to make:
		frame cleaning functions:
			-Input: takes in a 'dirty' dataframe
			-Outputs: a clean dataframe that can be more easily analyzed
		Notes:
			-Could create a function for each type of portfolio

TODO: Refactor to make portfolio cleaning functions one function
TODO: Figure out how to make a notes columns based on amount column values


Commands to add later:

	parser.add_argument('-prow', '--portfolio-rows', help='Enter the rows from the top of the portfolio file you want to skip', type=int)
	parser.add_argument('-pfooter', '--portfolio-footer-buyout', help='Enter the rows from the bottom to skip', type=int)
	parser.add_argument('-pname', '--portfolio-name', help='Enter the name of the portfolio', type=str)

'''

import pandas as pd
import numpy as np

# Functions for the project

DEFAULTCOLS = ('ContractNumber', 'CustomerName', 'Type', 'Bank Code', 'Amount', 'Program')
DEFAULTEXCEL = 'A,B,C,D,G,H'
DTYPES = {'Bank Code': np.object}
BUYOUTCOLUMNNAMES = ('ContractNumber', 'CustomerName', 'Amount', 'Note')
DEFAULTPORTFOLIOCOLUMNS = ('ContractNumber', 'CustomerName', 'Amount')
DEFAULTPORTFOLIOEXCELCOLUMNS = 'A, C, D'


def clean_ach(file, cols=DEFAULTCOLS, excelcolumns=DEFAULTEXCEL):
    ''' Cleans the ACH spreadsheet to grab the values needed for anlysis '''
    df = pd.read_excel(file, usecols=excelcolumns, names=cols, dtype=DTYPES)
    df = df.loc[(df['Type'] == 'U') & (df['Bank Code'] == '999.99'), :]
    df.reset_index(inplace=True)
    df.drop(columns=['index', 'Bank Code'], inplace=True)
    return df


def clean_portfolio(file, cut, rows=0, buyoutcolumns=BUYOUTCOLUMNNAMES):
    ''' Cleans portfolio speradsheet that has buyouts '''
    df = pd.read_excel(file, rows)
    buyouts = df.loc[df.ContractNumber == 'Buyout', :].copy()
    df.dropna(axis=0, thresh=1)
    df = df.iloc[:cut, :]
    df.dropna(axis=1, inplace=True)
    buyouts = buyouts[['CustomerName', 'Amount', 'Unnamed: 3', 'ContractNumber']]
    buyouts.columns = buyoutcolumns
    buyouts.Amount = np.round(buyouts.Amount.astype(float), 2)
    newdf = df.merge(buyouts, on='ContractNumber', how='left')
    newdf.loc[newdf.Note == 'Buyout', 'Amount_x'] = newdf.Amount_y
    newdf.drop(columns=['CustomerName_y', 'Amount_y'], inplace=True)
    newdf.fillna('', inplace=True)
    newdf.rename(columns={'Amount_x': 'Amount', 'CustomerName_x': 'Customer'}, inplace=True)
    return newdf


def clean_portfolio2(file, rows=2, footer=28):
    df = pd.read_excel(file, usecols=DEFAULTPORTFOLIOEXCELCOLUMNS, skip_footer=footer, skiprows=rows, names=DEFAULTPORTFOLIOCOLUMNS)
    df.ContractNumber.iloc[:9] = df.ContractNumber.apply(add_extension, args=('001',))
    df.ContractNumber.iloc[9:] = df.ContractNumber.apply(add_extension, args=('040',))
    df.dropna(axis=0, inplace=True)
    df = df.loc[df.CustomerName != 'Grand Total']
    df.loc[(df.Amount == 'Cancelled') | (df.Amount == 'Replaced'), 'Amount'] = 0
    df.Amount = df.Amount.astype(float)
    return df


def make_final_df(ach, port_df):
    final_df = ach.merge(port_df, on='ContractNumber', how='inner')
    final_df.drop(columns=['Program', 'CustomerName_x'], inplace=True)
    final_df.rename(columns={'Amount_x': 'Clarke LLC'}, inplace=True)
    final_df.insert(loc=5, column='Difference', value=final_df['Amount'] - final_df['Clarke LLC'])
    final_df = final_df.append(final_df.sum(numeric_only=True), ignore_index=True)
    final_df.fillna('', inplace=True)
    return final_df


def add_extension(contract, extension):
    if len(contract) == 11:
        return '-'.join([extension, contract])
    else:
        return contract


def main():
    DEFAULTCOLUMNS = ('ContractNumber', 'CustomerName', 'Type', 'Amount', 'Program')
    DEFAULTEXCEL = 'A,B,C,F,G'
    BUYOUTCOLUMNNAMES = ('ContractNumber', 'CustomerName', 'Amount', 'Note')
    try:
        ach_df = clean_ach('ach_test.xlsx', DEFAULTCOLUMNS, DEFAULTEXCEL, rows=20, footer=64)
        portfolio_df = clean_portfolio('portfolio.xlsx', 9, buyoutcolumns=BUYOUTCOLUMNNAMES)
        final_df = make_final_df(ach_df, portfolio_df)
        final_df.to_excel('Portfolio.xlsx')
    except FileNotFoundError as e:
        raise e
