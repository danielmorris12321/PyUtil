from re import sub
import os
import datetime
import pandas as pd
import dateutil
import xlwings as xw
import sys
import glob
import numpy as np
import shutil

import dm_exceptions as d_excep

from dateutil.relativedelta import relativedelta

# Dates
def return_date_formatted(s_date, curr_format, req_format):
    """
    Return a string date in a start format in a new format.
    """
    
    try: 
        curr_date = datetime.datetime.strptime(s_date, curr_format).date()
        reformat_date = curr_date.strftime(req_format)
        return reformat_date
    except:
        print("Error in converting " + s_date + " to " + req_format)

def return_datetime_obj(s_date, s_date_format, EOM = True):
    """
    Returns a datetime object from the string date. Used to set a value in excel equal to a date.
    """
    try:
        if EOM:
            dmonth = relativedelta(months =+ 1)
            dday = relativedelta(day =+ 1)
            current = datetime.datetime.strptime(s_date, s_date_format)
            current = current + dmonth
            current = current - dday
            value = current
        else:
            value = datetime.datetime.strptime(s_date, s_date_format)
        return value
    except:
        print("Error in converting " + s_date + " to datetime object.")

def roll_back(s_date, date_format = "%Y%m"):
    d_month = relativedelta(months =+ 1)
    date_1 = datetime.datetime.strptime(s_date, date_format).date()
    date_1 -= d_month
    date_1 = date_1.strftime(date_format)
    return date_1        

def copy_tree(src, dest):
    """
    Copy the directory tree from one base path to another. Doesn't return anything
    """
    try:
        shutil.copytree(src, dest)
    except shutil.Error as e:
        print("Directory not copied. Error: {}".format(e))
    

def dir_roll_forward(basepath, val_date, date_format= "%Y%m"):
    # TODO Make this smarter ...
    curr_path = regex_sub("YYYYMM", val_date, basepath)
    prev_date = roll_back(val_date, date_format)
    prev_path = regex.sub("YYYYMM", prev_date, basepath)
    copy_tree(prev_path, curr_path)
    print(prev_path + "has been rolled forward")

def create_dir(fpath, exist_ok = True):
    """
    Pass in a file path to have the folder directory created
    """
    try:
        direc = get_dir(fpath)
        os.makedirs(direc, exist_ok = exist_ok)
        print("Directory created: " + direc)
    except:
        raise d_excep.direcError(direc)

def get_dir(fpath):
    """
    Will return the folder path of a file passed in
    """
    get_dir = os.path.dirname(fpath)
    return get_dir

def get_file(fpath, include_ext = False):
    base = os.path.basename(fpath)
    if(include_ext):
        return base
    else:
        base = os.path.splitext(base)
        return base[0]

def find_pattern_file(direc, pattern, exclude = None):

    #TODO  -make this smarter

    direc = glob.glob(direc + str(pattern))
    filenames = [fn for fn in direc if not get_file(fn).startswith("~")]
    if exclude == None:
        for name in filenames:
            return name
    else:
        for name in filenames:
            if(exclude not in name):
                return name


def find_value(xl_sheet, target_val, row_max, col_max, row_min = 1, col_min = 1):
    for row in range(row_min, row_max):
        for col in range(col_min, col_max):
            if xl_sheet.range((row, col)).value == target_val:
                return xl_sheet.range((row, col)).address
    print("Unable to find " + str(target_val) + " in ")     

def init_xl_app(b_visible  = False, b_add_book = False, b_display_alerts = False):
    xl_app = xw.App(visible = b_visible, add_book = b_add_book)
    xl_app.display_alerts = b_display_alerts
    return xl_app

def close_xl_app(xl_app):
    books = xl_app.books()
    xl_app.quit()
    return xl_app

def xl_select_range(xl_sheet, start_range, expand_mode = "table"):
    xl_select_range = xl_sheet.range(start_range).expand(expand_mode)
    return xl_select_range

def clear_xl_range(xl_sheet, start_range, expand_mode = "None"):
    """
    Clears the contents of an excel range. Returns the sheet with the cleared down range.
    """

    if (expand_mode == "None"):
        curr_range = xl_sheet.range(start_range)
    else:
        curr_range = xl_select_range(xl_sheet, start_range, expand_mode)
    curr_range.clear_contents()
    return xl_sheet

def paste_to_range(xl_sheet, start_range, paste_data):
    """
    Paste data to a specific range in an excel sheet. Returns the excel sheet.
    """
    xl_sheet.range(start_range).value = paste_data
    return xl_sheet

def macro(xl_book, code_loc, macro_name):
    """
    Runs macro in an open xlbook. Specify the code_loc (eg 'Module1' and the macro name)
    TODO - make this smarter to take in parameters
    """
    macro_str = join_string([code_loc, macro_name], "!")
    xl_book.macro(macro_str)
    return(xl_book)

def clear_range_and_paste(xl_sheet, start_range, paste_data, expand_mode = "None"):
    xl_sheet = clear_xl_range(xl_sheet, start_range, expand_mode)
    xl_sheet = paste_to_range(xl_sheet, start_range, paste_data)
    return xl_sheet

def pd_first_col_to_header(df):
    """
    Simple function to move the first row in the dataframe to the column names
    """
    df.columns = df.iloc[0]
    df = df[1:]
    return df

def range_to_pd(xl_sheet, start_range, expand_mode = "table", first_row_header = True):
    xl_sheet = xl_sheet_filter(xl_sheet, auto_filter_mode = False)
    xl_data = xl_sheet.range(start_range).expand(expand_mode).value
    oput = pd.DataFrame(xl_data)
    if first_row_header:
        oput = pd_first_col_to_header(oput)

    return oput

def xl_sheet_filter(xl_sheet, auto_filter_mode = False):
    xl_sheet.api.AutoFilterMode = auto_filter_mode
    return xl_sheet

#CSV utils
 
def csv_to_pd():
    print(1)
    #TODO

# Regex utils
 
def regex_sub(search_str, new_string, base_string):
    x = sub(search_str, new_string, base_string)
    return x
 
def match_items_in_list(l_target, s_search):
    matching = [s for s in l_target if s_search in s]
    return matching

 
def match_folder_partial_string(base_path, pattern):
    # TODO make smarter - want one match or multiple
    matching_folder = glob.glob(base_path + pattern)
    return matching_folder

# Pandas utils
 
def df_clear_zero_cols(df):
    df = df.loc[:, (df != 0).any(axis = 0)]
    return df
 
def unique_in_col(df_data, col_name):
    #eg. col_name = "COL_1"
    oput  = df_data.filter([col_name]).squeeze().unique()
    return oput
 
def df_filter(df_data, filter_id, filter_col, col_keep = "All"):
# Pass in col_keep as a list
    if(col_keep == "All"):
        col_keep = df_get_colnames(df_data)
        
    try:
        df_data[df_data[filter_col].isin(filter_id)].filter(col_keep)
    except:
        print("Error in df_filter")

 
def df_get_colnames(df_data):
    cols = df_data.columns
    return(cols)

 
def unique_filter_results(df_data, col_keep, filter_id, filter_col):
    df_data = df_data[df_data[filter_col].isin(filter_id)]
    return df_data

#string utils 
def join_string(strings_target, s_sep = ""):
    new_str = s_sep.join(strings_target)
    return new_str

# list utils
def sort_list(l_data, reverse = False):
    """
    Converts a list into either ascending or descending order
    """
    l_data = sorted(l_data, reverse = reverse)
    return l_data

def remove_from_list(l_data, str_remove):
    while(str_remove in l_data):
        l_data.remove(str_remove)
    return l_data

def remove_duplicates_list(l_data):
    """
    Converts a list to a list of unique values
    """
    l_data = list(dict.fromkeys(l_data))
    return l_data