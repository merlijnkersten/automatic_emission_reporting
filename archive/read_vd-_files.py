import pandas as pd
import linecache

path = "C:/Veda/GAMS_WrkTIMES/nkep_250_upd_nuc010_0203.vd"

def file_extension(path):
    return path.split('.')[1]

def read_vd(path):
    # Read line 5 from the Veda file
    line4_str = linecache.getline(path, 4)

    # Get second part of the string, which contains the column names
    column_str = line4_str.replace('\n','').split('- ')[1]

    # Turn column_str into a list of column names.
    column_lst = column_str.split(';')

    # Read full data set (from line 13 onwards) without headers and setting correct column names
    return pd.read_csv(path, header=None, skiprows=12, names=column_lst)


def read_vde(path):
    # Columns list
    column_lst = ['COLUMN_A', 'Region', 'COLUMN_B', 'Discription']

    # Read full data set without headers and setting correct column names with ANSI codec
    return pd.read_csv(path, header=None, names=column_lst, encoding='ANSI')


def read_vds(path):
    # Columns list
    column_lst = ['COLUMN_A', 'Region', 'COLUMN_B', 'COLUMN_C']

    # Read full data set without headers and setting correct column names
    return pd.read_csv(path, header=None, names=column_lst)


def read_vdt(path):
    # Columns list
    column_lst = ['Region', 'COLUMN_A', 'COLUMN_B', 'IN_OUT']

    # Read full data set without headers and setting correct column names
    return pd.read_csv(path, header=None, skiprows=3, names=column_lst)

def read_VEDA_file(path):
    ext = file_extension(path)
    if ext == 'vd':
        data = read_vd(path)
    elif ext == 'vde':
        data = read_vde(path)
    elif ext == 'vds':
        data = read_vds(path)
    elif ext == 'vdt':
        data = read_vdt(path)
    else:
        raise ValueError(f"extension '{ext}' not in 'vd', 'vde', 'vds', 'vdt'")
    return data

for i in ["C:/Veda/GAMS_WrkTIMES/nkep_250_upd_nuc010_0203.vd",
          "C:/Veda/GAMS_WrkTIMES/nkep_250_upd_nuc010_0203.vde",
          "C:/Veda/GAMS_WrkTIMES/nkep_250_upd_nuc010_0203.vds",
          "C:/Veda/GAMS_WrkTIMES/cprice_090_with_mofo_2103.vdt",
          "C:/Veda/GAMS_WrkTIMES/nkep_250_upd_nuc010_0203.vdx"
         ]:
    print(f'\nfile: {i}')
    data = read_VEDA_file(i)
    print(data.head())

# Look up how to write tests
