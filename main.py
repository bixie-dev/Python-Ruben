import pandas as pd
from pandas import DataFrame
import os
import datetime

# [1] Constants
WHOLE = 'whole'  # Used for searching source files

# Destination Mode
FIXED = 'fixed'
INTEGER = 'integer'

# Percent Status
INCREASE = 'increase'
DECREASE = 'decrease'

# How to process
CLEAN_SRC = 'cleansourc'
CLEAN_DES = 'cleandest'
SEARCH_SRC_TO_DES = 'full - search from source to destination files'
SEARCH_DES_TO_SRC = 'full - search from destination to source files'
UPDATE_SRC_TO_DES = 'update inventory source files to destination files'
UPDATE_DES_TO_SRC = 'update inventory destination files to source files'
# [1]


# [2] Variables
# [2.1] Source files
src_routes = ['./sources']  # Source file routes (csv, xlv, xlsx)

# Column containing value for inventory from sources, if is 'whole', search whole columns
src_stock_columns = 'R'     # Stock value
src_sku_columns = 'O'       # SKU search
src_sheet_name = 'book2'    # Sheet name to work for source files
src_row_start = 0           # Beginning of content, not including header
# [2.1]

# [2.2] Destination files
des_routes = ['./destinations']  # Destination file routes
des_stock_columns = 'R'         # Stock value
des_sku_columns = 'O'           # SKU search
des_sheet_name = 'book2'        # Sheet name in destination files
des_row_start = 0               # Beginning of content, not including header
# [2.2]

min_stock = 50  # Minimun stock to search in des_routes

safe_files = ['safe.xlsx']
safe_files_columns = 'B,D'
safe_files_enabled = False

# If true, will remove rows with lower stock value than entered in source files
src_stock_filter_enabled = True

# If true, will remove rows with lower stock value than entered in destination files
des_stock_filter_enabled = True

# How destination columns saved in destination files
des_mode = FIXED
if (des_mode == FIXED):
    des_decimals = 2                # Decimal counts

# If true, it will increase or decrease a percent from 'percent_status'
is_percent_added_to_dest = True
if (is_percent_added_to_dest):
    percent = 5
    percent_status = INCREASE

rows_removable = True  # If true, it will remove rows from activities

output_route = './outputs'
report_route = './reports'

fill_zeros = True

# how_to_process = CLEAN_SRC
# how_to_process = CLEAN_DES
# how_to_process = SEARCH_SRC_TO_DES
# how_to_process = SEARCH_DES_TO_SRC
how_to_process = UPDATE_SRC_TO_DES
# how_to_process = UPDATE_DES_TO_SRC

# [2]


# [3] Functions
def no_wsp(string: str) -> str:  # Return string which is removed whitespaces
    return string.replace(' ', '')


# Convert given number to an Excel-style column name
def column_in_excel(col: int) -> str:
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    result = []
    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result)


def read_folder(folder: str) -> list[str]:
    return os.listdir(folder)


def read_file(file_name: str, file_type: str, sheet_name: str, start: int):
    data_frame = None
    name = (file_name + file_type).strip()

    if os.path.exists(name):
        if file_type == '.xlsx' or file_type == '.xls':
            data_frame = pd.read_excel(name, sheet_name=sheet_name,
                                       skiprows=start)
            init_columns = data_frame.columns
            number_of_columns = len(data_frame.columns)
            data_frame.columns = [column_in_excel(i)
                                  for i in range(1, number_of_columns + 1)]

        elif file_type == '.csv':  # reading source file with its sheet
            data_frame = pd.read_csv(name, sep=',', lineterminator='\r',
                                     sheet_name=sheet_name, skiprows=start)
            number_of_columns = len(data_frame.columns)
            init_columns = data_frame.columns
            data_frame.columns = [column_in_excel(i)
                                  for i in range(1, number_of_columns+1)]

        elif file_type == '.txt':  # reading source file with its sheet
            data_frame = pd.read_csv(name, sep='\t', lineterminator='\r',
                                     skiprows=start)
            number_of_columns = len(data_frame.columns)
            init_columns = data_frame.columns
            data_frame.columns = [column_in_excel(i)
                                  for i in range(1, number_of_columns+1)]
            data_frame = data_frame.replace(r'\n', '', regex=True)
    else:
        print(file_name + " does not exist.")

    return (data_frame, init_columns)


def save_to_report(data: DataFrame, file_name: str, file_extension: str):
    if not os.path.exists(report_route):
        os.makedirs(report_route)

    if file_extension.lower() == '.xlsx':
        data.to_excel(report_route + "/" + file_name + ".xlsx", index=False)
    elif file_extension.lower() == '.csv':
        data.to_csv(report_route + "/" + file_name + ".csv", index=False)
    else:
        data.to_csv(report_route + "/" + file_name +
                    ".txt", sep="\t", index=False)


def save_to_output(data: DataFrame, file_name: str, file_extension: str):
    if not os.path.exists(output_route):
        os.makedirs(output_route)

    if file_extension.lower() == '.xlsx':
        data.to_excel(output_route+"/"+file_name+".xlsx", index=False)

    elif file_extension.lower() == '.csv':
        data.to_csv(output_route+"/"+file_name+".csv", index=False)

    else:
        data.to_csv(output_route+"/"+file_name+".txt", sep="\t", index=False)


def clean(routes: list[str], sheet_name: str, row_start: int, stock_columns: str):
    for route in routes:
        files = read_folder(route)

        for file in files:
            file_name, file_extension = os.path.splitext(route + '/' + file)
            df, init_columns = read_file(file_name, file_extension,
                                         sheet_name, row_start)

            print('processing ' + file_name + file_extension + ' . . .')
            print(df)

            if df is not None:
                columns = (no_wsp(des_stock_columns).split(','),
                           df.columns)[stock_columns.lower() == WHOLE]

                df_temp = df
                for column in columns:
                    df_temp = df_temp[df_temp[column] > min_stock]

                removed_values = pd.concat([df, df_temp]) \
                    .drop_duplicates(keep=False)
                removed = len(removed_values) > 0 and 'removed' \
                    or 'not-removed'
                removed_values.columns = df_temp.columns = init_columns

                save_to_output(df_temp,
                               file + '~' + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                               file_extension)

                save_to_report(removed_values,
                               file + '~' + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                               file_extension)

                print("Procssing " + file_name + file_extension
                      + " is complete &  saved to " + file)


# Search from a to b
def search(a_routes: list[str], b_routes: list[str], a_sheet_name: str, b_sheet_name: str,
           a_row_start: int, b_row_start: int, a_sku_columns: str, b_sku_columns: str,
           b_stock_columns: str):
    for a_route in a_routes:
        a_files = read_folder(a_route)
        for a_file in a_files:
            a_file_name, a_file_extension = os.path.splitext(a_route
                                                             + '/' + a_file)
            a_df, init_columns = read_file(a_file_name, a_file_extension,
                                           a_sheet_name, a_row_start)
            print('processing ' + a_file_name +
                  a_file_extension + ' . . . with')

            for b_route in b_routes:
                b_files = read_folder(b_route)
                for b_file in b_files:
                    b_file_name, b_file_extension = os.path.splitext(b_route + '/'
                                                                     + b_file)
                    b_df, init_columns = read_file(b_file_name, b_file_extension,
                                                   b_sheet_name, b_row_start)
                    print('  processing ' + b_file_name + b_file_extension)

                    if a_df is not None and b_df is not None:
                        a_sku_column_list = (no_wsp(a_sku_columns).split(','),
                                             a_df.columns)[a_sku_columns == WHOLE]
                        b_sku_column_list = (no_wsp(b_sku_columns).split(','),
                                             b_df.columns)[b_sku_columns == WHOLE]

                        b_stock_column_list = (no_wsp(b_stock_columns).split(','),
                                               b_df.columns)[b_stock_columns == WHOLE]

                        filtered_df = b_df
                        found_df = DataFrame()

                        for a_column in a_sku_column_list:
                            for b_column in b_sku_column_list:
                                filtered_df = filtered_df[filtered_df[b_column]
                                                          .isin(a_df[a_column].values) == True]
                                found_df = pd.concat([found_df,
                                                      filtered_df[filtered_df[b_column].isin(a_df[a_column]) == True]])

                        filtered_df = found_df.drop_duplicates()
                        found_df = DataFrame()

                        print(filtered_df)
                        for b_column in b_stock_column_list:
                            filtered_df = filtered_df[filtered_df[b_column]
                                                      > min_stock]
                            found_df = pd.concat([found_df,
                                                  filtered_df])

                        found_df = found_df.drop_duplicates()
                        removed_df = pd.concat([b_df, found_df]) \
                            .drop_duplicates(keep=False)

                        removed = 'removed' if len(removed_df) > 0 \
                            else 'not-removed'
                        found_df.columns = removed_df.columns = init_columns

                        save_to_output(found_df,
                                       a_file + '~' + b_file + '~' + removed + '~'
                                       + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                                       b_file_extension)
                        save_to_report(removed_df,
                                       a_file + '~' + b_file + '~' + removed + '~'
                                       + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                                       b_file_extension)
                        print("Procssing " + a_file_name + a_file_extension +
                              "with" + b_file_name + b_file_extension + " is complete")


def update(a_routes: list[str], b_routes: list[str], a_sheet_name: str, b_sheet_name: str,
           a_row_start: int, b_row_start: int, a_sku_columns: str, b_sku_columns: str,
           a_stock_columns: str, b_stock_columns: str):
    for a_route in a_routes:
        a_files = read_folder(a_route)
        for a_file in a_files:
            a_file_name, a_file_extension = os.path.splitext(a_route
                                                             + '/' + a_file)

            a_df, init_columns = read_file(a_file_name, a_file_extension,
                                           a_sheet_name, a_row_start)
            print('processing ' + a_file_name +
                  a_file_extension + ' . . . with')

            for b_route in b_routes:
                b_files = read_folder(b_route)
                for b_file in b_files:
                    b_file_name, b_file_extension = os.path.splitext(b_route + '/'
                                                                     + b_file)
                    b_df, init_columns = read_file(b_file_name, b_file_extension,
                                                   b_sheet_name, b_row_start)

                    print('processing ' + b_file_name + b_file_extension)

                    if a_df is not None and b_df is not None:
                        a_sku_column_list = (no_wsp(a_sku_columns).split(','),
                                             a_df.columns)[a_sku_columns == WHOLE]
                        b_sku_column_list = (no_wsp(b_sku_columns).split(','),
                                             b_df.columns)[b_sku_columns == WHOLE]

                        a_stock_column_list = (no_wsp(a_stock_columns).split(','),
                                               a_df.columns)[a_stock_columns == WHOLE]
                        b_stock_column_list = (no_wsp(b_stock_columns).split(','),
                                               b_df.columns)[b_stock_columns == WHOLE]

                        filtered_df = b_df
                        found_df = DataFrame()

                        for a_column in a_sku_column_list:
                            for b_column in b_sku_column_list:
                                filtered_df = filtered_df[filtered_df[b_column]
                                                          .isin(a_df[a_column].values) == True]
                                found_df = pd.concat([found_df,
                                                      filtered_df[filtered_df[b_column].isin(a_df[a_column]) == True]])

                        filtered_df = found_df.drop_duplicates()
                        found_df = DataFrame()

                        print(filtered_df)
                        for b_column in b_stock_column_list:
                            filtered_df = filtered_df[filtered_df[b_column]
                                                      > min_stock]
                            found_df = pd.concat([found_df,
                                                  filtered_df])

                        found_df = found_df.drop_duplicates()
                        removed_df = pd.concat([b_df, found_df]) \
                            .drop_duplicates(keep=False)

                        for m, n in zip(b_sku_column_list, b_stock_column_list):
                            v = found_df[m]
                            print(v)
                            for index in v:
                                for i, j in zip(a_sku_column_list, a_stock_column_list):
                                    x = a_df[a_df[i] == index][j]
                                    y = found_df[found_df[m]
                                                 == index][n]
                                    found_df \
                                        .loc[found_df[m] == index, n] = float(x)

                        if fill_zeros:
                            for k in b_stock_column_list:
                                removed_df[k] = 0
                            found_df = pd.concat([found_df,
                                                  removed_df])

                        removed = 'removed' if len(removed_df) > 0 \
                            else 'not-removed'
                        found_df.columns = removed_df.columns = init_columns

                        save_to_output(found_df,
                                       a_file + '~' + b_file + '~' + removed + '~'
                                       + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                                       b_file_extension)
                        save_to_report(removed_df,
                                       a_file + '~' + b_file + '~' + removed + '~'
                                       + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
                                       b_file_extension)
                        print("Procssing " + a_file_name + a_file_extension
                              + "with" + b_file_name + b_file_extension + " is complete")
# [3]


# [4] Main function
def main():
    if (how_to_process == CLEAN_SRC):
        clean(src_routes, src_sheet_name, src_row_start, src_stock_columns)

    if (how_to_process == CLEAN_DES):
        clean(des_routes, des_sheet_name, des_row_start, des_stock_columns)

    if (how_to_process == SEARCH_SRC_TO_DES):
        search(src_routes, des_routes, src_sheet_name, des_sheet_name, src_row_start, des_row_start,
               src_sku_columns, des_sku_columns, des_stock_columns)

    if (how_to_process == SEARCH_DES_TO_SRC):
        search(des_routes, src_routes, des_sheet_name, src_sheet_name, des_row_start, src_row_start,
               des_sku_columns, src_sku_columns, src_stock_columns)

    if (how_to_process == UPDATE_SRC_TO_DES):
        update(src_routes, des_routes, src_sheet_name, des_sheet_name, src_row_start, des_row_start,
               src_sku_columns, des_sku_columns, src_stock_columns, des_stock_columns)

    if (how_to_process == UPDATE_DES_TO_SRC):
        update(des_routes, src_routes, des_sheet_name, src_sheet_name, des_row_start, src_row_start,
               des_sku_columns, src_sku_columns, des_stock_columns, src_stock_columns)


if __name__ == "__main__":
    main()
# [4]
