from portfolio_tie_out import clean_ach, clean_portfolio, make_final_df

# ACH function tests


def test_default_ach_dataframe_shape():
    # Tests that ach function is returning the proper values from the excel file
    df = clean_ach('ach_test.xlsx')
    assert df.shape == (200, 5)


def test_default_ach_dataframe_columns():
    # Tests that the columns of the dataframe are correct
    df = clean_ach('ach_test.xlsx')
    assert list(df.columns) == ['ContractNumber', 'CustomerName', 'Type', 'Amount', 'Program']


def test_ach_dataframe_dtypes():
    # Test to make sure the dtypes are correct
    df = clean_ach('ach_test.xlsx')
    assert all(df.dtypes == ['O', 'O', 'O', 'float64', 'O'])


def test_ach_dataframe_n_types():
    # Tests if there is any N option ach amounts under the Type column
    df = clean_ach('ach_test.xlsx')
    assert 'N' not in df.Type

# Dataframe merge tests


def test_final_dataframe_columns():
    df_ach = clean_ach('ach_test.xlsx')
    df_portfolio = clean_portfolio('portfolio.xlsx', 'Portfolio_tie_out060218.xlsx', buyout=True, cut=9)
    df_final = make_final_df(df_ach, df_portfolio, portfolio_name='Clarke LLC')
    assert list(df_final.columns) == ['ContractNumber', 'CustomerName', 'Type', 'Amount', 'Clarke LLC', 'Difference', 'Notes']


def test_final_dataframe_shape():
    df_ach = clean_ach('ach_test.xlsx')
    df_portfolio = clean_portfolio('portfolio.xlsx', 'Portfolio_tie_out060218.xlsx', buyout=True, cut=9)
    df_final = make_final_df(df_ach, df_portfolio, portfolio_name='Clarke LLC')
    assert df_final.shape == (10, 7)


def test_final_dataframe_dtypes():
    df_ach = clean_ach('ach_test.xlsx')
    df_portfolio = clean_portfolio('portfolio.xlsx', 'Portfolio_tie_out060218.xlsx', buyout=True, cut=9)
    df_final = make_final_df(df_ach, df_portfolio, portfolio_name='Clark LLC')
    assert all(df_final.dtypes == ['O', 'O', 'O', 'float64', 'float64', 'float64', 'O'])
