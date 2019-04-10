import openpyxl
import pandas as pd

def census_by_county(fn_xlsx, sheet_name, fn_csv):

    """
    Tabulates population and number of census tracts for each county.

    Args:
        fn_xlsx:    filename of Excel file to process
        sheet_name: sheet name within Excel file to be processed
        fn_csv:     filename of output CSV file
    Returns:
        nothing, just writes CSV file

    """
    
    # load workbook and select sheet of interest
    wb = openpyxl.load_workbook(fn_xlsx)
    ws = wb[sheet_name]

    # create data frame
    data = ws.values
    header = next(data)[0:]
    df = pd.DataFrame(data, columns = header)

    # create grouped data frame
    df2 = df[['State','County','POP2010']].groupby(by = ['State', 'County']).agg(['sum', 'count'])
    df2.reset_index(inplace = True)
    df2.columns = ['State','County','Census_Tract', 'Population']
    
    # write csv file
    df2.to_csv(fn_csv)
    
    

if __name__ == '__main__':

    xlsxfile = 'data/censuspopdata.xlsx'
    csvfile = 'census_sorted.csv'
    sheetname = 'Population by Census Tract'

    census_by_county(xlsxfile, sheetname, csvfile)