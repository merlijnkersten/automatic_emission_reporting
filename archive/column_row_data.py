
from distutils.log import error


def number_to_letter(i):
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    #           12345678901234567890123456
    if i in range(0,25):
        return alphabet[i]
    else:
        error(f"{i} out of range (0-25)")

ghg_col_dct = {
    'NOx' : {
        'historic' : 'C',
        '2020'     : 'D',
        '2025'     : 'E',
        '2030'     : 'F',
        '2040'     : 'G',
        '2050'     : 'H'
    },

    'NMVOC' : {
        'historic' : 'I',
        '2020'     : 'J',
        '2025'     : 'K',
        '2030'     : 'L',
        '2040'     : 'M',
        '2050'     : 'N'
    },

    'SOx' : {
        'historic' : 'O',
        '2020'     : 'P',
        '2025'     : 'Q',
        '2030'     : 'R',
        '2040'     : 'S',
        '2050'     : 'T'
    },

    'NH3' : {
        'historic' : 'U',
        '2020'     : 'V',
        '2025'     : 'W',
        '2030'     : 'X',
        '2040'     : 'Y',
        '2050'     : 'Z'
    },

    'PM2_5' : {
        'historic' : 'AA',
        '2020'     : 'AB',
        '2025'     : 'AC',
        '2030'     : 'AD',
        '2040'     : 'AE',
        '2050'     : 'AF'
    },
}

def ghg_col(ghg, period, dct=ghg_col_dct):
    return dct[ghg][period]

# Created from a dictionary (zipped) from columns NFR Code and Longname in the Excel file
nfr_short_long_name = {
    '1a1': 'Energy industries (Combustion in power plants & Energy Production)',
    '1a2': 'Manufacturing Industries and Construction (Combustion in industry including Mobile)',
    '1a3b': 'Road Transport', 
    '1a3bi': 'R.T., Passenger cars', 
    '1a3bii': 'R.T., Light duty vehicles', 
    '1a3biii': 'R.T., Heavy duty vehicles', 
    '1a3biv': 'R.T., Mopeds & Motorcycles', 
    '1a3bv': 'R.T., Gasoline evaporation', 
    '1a3bvi': 'R.T., Automobile tyre and brake wear', 
    '1a3bvii': 'R.T., Automobile road abrasion', 
    '1a3a_c_d_e': 'Off-road transport', 
    '1a4': 'Other sectors (Commercial, institutional, residential, agriculture and fishing stationary and mobile combustion)', 
    '1a5': 'Other', 
    '1a': 'Fugitive emissions (Fugitive emissions from fuels)', 
    '2a_b_c_h_i_j_k_l': 'Industrial Processes', 
    '2d_2g': 'Solvent and other product use', 
    '3b': 'Animal husbandry and manure management', 
    '3b1a': 'Cattle Dairy', '3B1b': 'Cattle Non-Dairy', 
    '3b2': 'Sheep', 
    '3b3': 'Swine', 
    '3b4a': 'Buffalo', 
    '3b4d': 'Goats', 
    '3b4e': 'Horses', 
    '3b4f': 'Mules and asses', 
    '3b4g': 'Poultry', 
    '3b4h': 'Other', 
    '3d': 'Plant production and agricultural soils', 
    '3f_i': 'Field burning and other agriculture', 
    '5': 'Waste', 
    '6a': 'Other (included in National Total for Entire Territory)', 
    'nationaltotal': 'National Total for the entire territory'
}

def redo_name(i):
    return i.replace(',','_').replace(' ','').lower()

nfr_lst = [redo_name(i) for i in list(nfr_short_long_name.keys())]

# creates the following list:
# nfr_list = ['1a1', '1b2', '1a3b', '1a3bi', '1a3bii', '1a3biii', '1a3biv', '1a3bv', '1a3bvi',
#            '1a3bvii', '1a3a_c_d_e', '1a4', '1a5', '1a', '2a_b_c_h_i_j_k_l', '2d_2g', '3b', 
#            '3b1a', '3B1b', '3b2', '3b3', '3b4a', '3b4d', '3b4e', '3b4f', '3b4g', '3b4h', 
#            '3d', '3f_i', '5', '6a', 'nationaltotal']

def nfr_row(i, lst=nfr_lst):
    return str(lst.index(i) + 12)