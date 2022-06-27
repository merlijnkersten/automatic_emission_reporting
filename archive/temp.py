import pandas as pd

prices_paths = [f"C:/Users/Merlijn Kersten/Documents/Univerzita Karlova/Timeslice analysis/prices {year}.xls" for year in range(2015,2022)]

def generate_prices_csv(prices_paths):

    df_lst = []

    for path in prices_paths:
        df = pd.read_excel(path, skiprows=5,  engine='xlrd', sheet_name='DAM')

        df = df[['Day', 'Hour', 'Marginal price CZ (EUR/MWh)', 'Marginal price CZ (CZK/MWh)']]

        df = df[ df['Hour'] != 25 ]

        def zero_padded_hour(i):
            j = str(i-1)
            if len(j) == 1:
                return ' 0' + j
            else:
                return ' '  + j

        df['Date and time'] = df['Day'].dt.strftime(r'%Y-%m-%d') + df['Hour'].apply(zero_padded_hour)

        df['Date and time'] = pd.to_datetime(df['Date and time'], format=r'%Y-%m-%d %H')

        df.set_index("Date and time", inplace=True)

        df.drop(columns=['Day', 'Hour'], inplace=True)

        df.columns = ['Price [EUR/MWh]', 'Price [CZK/MWh]']

        df_lst.append(df)

    combined_df = pd.concat(df_lst)

    path = r"C:/Users/Merlijn Kersten/Documents/Univerzita Karlova/Timeslice analysis/prices 2015-2021.csv"

    combined_df.to_csv(path)

def fmt_col(col):
    for string in [' [MW]', ' [EUR/MWh]', ' [CZK/MWh]']:
        col = col.replace(string, '')
    return col.lower()


path = r"C:/Users/Merlijn Kersten/Documents/Univerzita Karlova/Timeslice analysis/prices 2015-2021.csv"

def read_prices(path, combined_df):
    price_df = pd.read_csv(path)

    price_df['Date and time'] = pd.to_datetime(price_df['Date and time'], format=r'%Y-%m-%d %H', )
    price_df.set_index('Date and time', inplace=True)

    df = pd.merge(price_df, combined_df, left_index=True, right_index=True)

    return df

df = read_prices(path)

print(df.head())

def import_prices(lst):
    '''
    lst: list of paths
    loads CSVs and combines them into a single, polished file.
    '''

    # Empty list, to be populated with DataFrames for every year.
    dfs_lst = []

    # Translation dictionary from TSO names to their respective countries
    tso_dct = {
        'PSE Actual [MW]'   : 'Poland [MW]',
        'SEPS Actual [MW]'  : 'Slovakia [MW]',
        'APG Actual [MW]'   : 'Austria [MW]',
        'TenneT Actual [MW]': 'Germany (south) [MW]',
        '50HzT Actual [MW]' : 'Germany (north) [MW]',
        'CEPS Actual [MW]'  : 'Net import [MW]'
    }

    # Unecessary columns, will be removed later.
    redundant_columns = ['Unnamed: 13', 'Date', 'PSE Planned [MW]', 'SEPS Planned [MW]', 
        'APG Planned [MW]', 'TenneT Planned [MW]', '50HzT Planned [MW]','CEPS Planned [MW]']


    for path in lst:
        # Read cross border data
        df = pd.read_csv(path, skiprows=2, sep=';')
        # Create a timestamp column based on "Date" and set this as the index
        df["Date and time"] = pd.to_datetime(df["Date"], format=r"%d.%m.%Y %H:%M")
        df.set_index("Date and time", inplace=True)       
        # Drop unecessary columns
        df.drop(columns=redundant_columns, inplace=True)
        # Rename columns from TSO names to country names
        df.columns = [tso_dct[i] for i in list(df.columns)]
        # Append the Dataframe to the Datafraemes list
        dfs_lst.append(df)

    # Concatenate dataframes together
    df = pd.concat(dfs_lst)

    # Create import and export columns using the find_import_export function
    df['Imports [MW]'] = df.apply(lambda i: find_import_export(i, 'import'), axis=1)
    df['Exports [MW]'] = df.apply(lambda i: find_import_export(i, 'export'), axis=1)

    return df