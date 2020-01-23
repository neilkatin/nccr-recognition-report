#! /usr/bin/env python3

# convert.py -- convert one spreadsheet into another

import logging
import os
import os.path
import re
import datetime
import time
import pprint
import json
import io
import csv
import random
import textwrap
import pathlib
import argparse

import requests
import requests_html
import dotenv
import xlsxwriter
import xlrd



from logging.config import dictConfig

def main():

    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        log.debug("Logging level set to debug")


    log.debug("in main")

    dotenv.load_dotenv()

    outfolder = pathlib.Path('.') / 'output';
    if not os.path.exists(outfolder):
        os.makedirs(outfolder)

    log.debug(f"output folder is '{ outfolder }'")

    infolder = pathlib.Path('.') / 'input';
    if not os.path.exists(infolder):
        log.fatal(f"input folder { infolder } does not exist")
        return(1)

    now = datetime.datetime.now()
    datestamp = now.strftime("%Y-%m-%d")

    make_report(outfolder, infolder, datestamp)


def make_report(outfolder, infolder, datestamp):

    infilename = infolder / f"input.xls"
    outfilename = infolder / f"{ datestamp }_Recognition-report.xlsx"

    wb = xlsxwriter.Workbook(outfilename)

    date_format = wb.add_format({ 'num_format': 'yyyy-mm-dd;@', 'align': 'left' })

    # read in the staff requests spreadsheet
    title_row, data_rows = read_input_document(infilename)
    column_map = title_row_to_dict(title_row)

    districts = extract_column_data(data_rows, column_map, 'District')

    log.debug(f"districts: { districts }")

    for district in districts:
        generate_sheet(wb, district, data_rows, column_map)

    return

    staff_data = read_staff_roster(staff_ws)

    copy_staff_roster(staff_data, ws0, date_format, gaps, None, lambda x: True)
    copy_staff_roster(staff_data, ws1, date_format, gaps, 'Reporting/Work Location', lambda x: x.startswith('MC/SH/'))
    copy_staff_roster(staff_data, ws2, date_format, gaps, 'GAP(s)', lambda x: x.startswith('MC/') and not x.startswith('MC/SH/'))
    copy_staff_roster(staff_data, ws3, date_format, gaps, 'GAP(s)', lambda x: not x.startswith('MC/'))
    copy_staff_roster(staff_data, ws7, date_format, gaps, None, lambda x: not x.endswith('/SA'))

    do_arrival_roster(ws4, folder / f"{datestamp}-ArrivalRoster.xls", date_format)
    do_open_staff_requests(ws5, folder / f"{datestamp}-OpenStaffRequests.xls", date_format)
    do_air_travel_roster(ws6, folder / f"{datestamp}-AirTravelRoster.xls", date_format)

    wb.close()

def generate_sheet(wb, district, data_rows, column_map):
    """ generate a new sheet with just that district in it """

    pass

def extract_column_data(data_rows, column_map, column_name):
    """ extract all the values of a given column.

        Values will be unique and sorted
    """

    if column_name not in column_map:
        raise Exception(f"Could not find column { column_name } in column_map")

    column_index = column_map[column_name]

    result_map = {}
    for row in data_rows:
        value = row[column_index]
        result_map[value] = 1

    return sorted(result_map.keys())




def read_input_document(filename):
    wb = xlrd.open_workbook(filename)
    ws = wb.sheet_by_index(0)

    title_row = ws.row_values(1)
    data_rows = []

    log.debug(f"title_row: { title_row }")

    for rownum in range(3, ws.nrows):
        value = ws.row_values(rownum)
        data_rows.append(value)

    log.debug(f"# rows { len(data_rows) }, ncols { ws.ncols }")

    return title_row, data_rows

def read_staff_roster(ws):

    title_row = ws.row_values(5)
    output_list, col_map = make_staff_copylist(title_row)

    data = []
    for rownum in range(5, ws.nrows):
        input_row = ws.row_values(rownum)
        output_row = []
        for colnum in range(0, ws.ncols):

            source_col = output_list[colnum]
            source_val = input_row[source_col]

            output_row.append(source_val)

        data.append(output_row)

    return data


def copy_staff_roster(data, destws, date_format, gaps, sort_column, gap_filter):

    date_columns = { 'Assigned': 1, 'Checked in': 1, 'Released': 1, 'Expect release': 1 }
    num_columns = { 'On Job': 1, 'DaysRemain': 1, 'Mem#': 1 }

    copy_sheet(data, destws, num_columns, date_columns, date_format, sort_column, 'GAP(s)', gap_filter)

    title_map = title_row_to_dict(data[0])
    gap_column = title_map['GAP(s)']

    fixup_staff_roster(data, destws)
    list_gaps = list(filter(gap_filter, gaps))
    destws.filter_column_list(gap_column, list_gaps)
 
def copy_sheet(data, destws, num_columns, date_columns, date_format, sort_column, filter_column, filter_func):

    #log.debug(f"title_row { data[0] }")

    title_map = title_row_to_dict(data[0])

    if sort_column != None:
        data = data.copy()
        title_row = data.pop(0)
        sort_index = title_map[sort_column]
        log.debug(f"copy_staff_roster: sorting on { sort_column } / { sort_index }")
        data.sort(key=lambda x: x[sort_index])
        data.insert(0, title_row)


    # ignore the first five rows, which are merged titles
    for rownum, row in enumerate(data):

        for colnum, source_val in enumerate(row):

            source_title = data[0][colnum]

            if rownum != 0 and source_title in num_columns:
                # convert to number
                if source_val != '':
                    destws.write_number(rownum, colnum, int(source_val))
            elif source_title in date_columns:
                # treat as a date
                destws.write(rownum, colnum, source_val, date_format)
            else:
                # default
                destws.write(rownum, colnum, source_val)

            if rownum != 0 and source_title == filter_column and filter_func != None:
                # hide column if its filtered out
                if not filter_func(source_val):
                    destws.set_row(rownum, options={'hidden': True})


def make_staff_copylist(row):


    preferred_cols = [ 'Name', 'GAP(s)', 'Cell phone', 'Reporting/Work Location' ]
    return make_copylist(row, preferred_cols)

def make_copylist(row, preferred_cols):
    """ make a list to copy rows from, moving rows around to match desired order """

    # basic strategy: copy named rows to the specified position, then put the rest in spreadsheet order
    #log.debug(f"row { row }")

    col_map = title_row_to_dict(row)
    output_col_map = {}
    #log.debug(f"dict { dict }")

    processed_cols = {}

    # do the preferred columns
    output_list = []
    for colname in preferred_cols:
        colnum = col_map[colname]
        output_col_map[colname] = len(output_list)
        output_list.append(colnum)
        processed_cols[colname] = True

    #do the rest
    for colname in row:
        if colname not in processed_cols:
            output_col_map[colname] = len(output_list)
            output_list.append(col_map[colname])

    #log.debug(f"output_list { output_list }")
    #log.debug(f"col_map { col_map }")

    return output_list, output_col_map



def title_row_to_dict(row):
    """ turn a title row into a dict of name -> column-number """

    dict = {}
    colnum = 0

    for val in row:
        dict[val] = colnum
        colnum += 1

    return dict


def fixup_staff_roster(data, sheet):
    """ do standard fixups on the sheet """

    hidden_columns = [ 'Mem#', 'Region', 'State', 'Assigned', 'Res', 'Last action', 'Ge', 'Released', 'Acc' ]
    width_columns = { 'Name': 30, 'Preferred name': 15, 'GAP(s)': 14, 'Cell phone': 13, 'Checked in': 12,
            'Released': 12, 'Expect release': 12,
            'Reporting/Work Location': 40, 'District': 20, 'Current lodging': 30,
            'Qualifications': 30, 'All GAPs': 50, 'Languages': 20, 'Email': 30, 'Home phone': 13, 'Work phone': 13
            }

    fixup_sheet(data, sheet, width_columns, hidden_columns)


def fixup_sheet(data, sheet, width_columns, hidden_columns):

    column_map = title_row_to_dict(data[0])
    nrows = len(data)
    ncols = len(data[0])

    for colname, width in width_columns.items():
        colnum = column_map[colname]

        #log.debug(f"setting column width of { colname } / { colnum } to { width }")
        sheet.set_column(colnum, colnum, width)

    for colname in hidden_columns:
        colnum = column_map[colname]

        #log.debug(f"Hiding column { colname } / { colnum }")
        sheet.set_column(colnum, colnum, None, None, { 'hidden': 1 })

    sheet.freeze_panes('B2')
    sheet.autofilter(0, 0, nrows-1, ncols -1)



def do_arrival_roster(destws, filename, date_format):

    #log.debug("do_arrival_roster: called")

    data = read_sheet(filename, 5, [ 'Name', 'GAP', 'Cell phone' ])
    copy_sheet(data, destws, [], [ 'Arrive date' ], date_format, None, None, None)

    width_columns = {
            'Name': 30, 'GAP': 14, 'Cell phone': 12,  'Arrive date': 11,
            'Reporting/Work Location': 30, 'Email': 20, 'Home phone': 12, 'Work phone': 12
            }
    hidden_columns = []

    fixup_sheet(data, destws, width_columns, hidden_columns)

    #log.debug("do_arrival_roster: done")

def do_open_staff_requests(destws, filename, date_format):

    data = read_sheet(filename, 1, [])
    copy_sheet(data, destws, [], [], date_format, None, None, None)

    width_columns = {
            'G/A/P': 14, 'Proximity': 14,
            }
    hidden_columns = []

    fixup_sheet(data, destws, width_columns, hidden_columns)


def do_air_travel_roster(destws, filename, date_format):

    data = read_sheet(filename, 3, ['Name', 'GAP', 'Cell Number'])
    copy_sheet(data, destws, [], ['Arrival Date/Time'], date_format, None, None, None)

    width_columns = {
            'Name': 30, 'GAP': 14, 'Cell Number': 12, 'Arrival Date/Time': 12,
            'Arrival City': 18, 'Departure City': 18, 'Airline': 18, 'Assign/CheckIn': 14,
            'Region name': 28, 'Status': 16,
            }
    hidden_columns = []

    fixup_sheet(data, destws, width_columns, hidden_columns)


def read_sheet(filename, skip_rows, preferred_columns):
    wb = xlrd.open_workbook(filename)
    ws = wb.sheet_by_index(0)

    title_row = ws.row_values(skip_rows)
    output_list, output_col_list = make_copylist(title_row, preferred_columns)

    # read the data
    data = []
    for rownum in range(skip_rows, ws.nrows):
        input_row = ws.row_values(rownum)
        output_row = []
        for colnum in range(0, ws.ncols):

            source_col = output_list[colnum]
            source_val = input_row[source_col]

            if title_row[source_col] == '':
                #log.debug(f"skipping col { source_col } because title col is empty")
                continue

            output_row.append(source_val)

        data.append(output_row)

    return data



def parse_args():

    description = textwrap.dedent("""\

            Convert a recognition report from volunteer-connection report format to the one split by District

            """)

    parser = argparse.ArgumentParser(
            formatter_class=argparse.RawDescriptionHelpFormatter,
            description=description)

    parser.add_argument('--debug', action='store_true', help=f"Turn on additional debug output")

    args = parser.parse_args()

    return args


def init_logging(app_name):
    logging_config = {
        'version': 1,
        'handlers': {
            'console': {
                'class': 'logging.StreamHandler',
                'formatter': 'default',
                'level': 'DEBUG',
                'stream': 'ext://sys.stderr'
            },
        },
        'formatters': {
            'default': {
                'format': '%(asctime)s %(levelname)-5s %(name)-10s %(funcName)-15.15s:%(lineno)4d %(message)s',
            'datefmt': '%Y-%m-%d %H:%M:%S',
            },
        },
        'root': {
            'level': 'INFO',
            'handlers': [ 'console' ],
        },
        'loggers': {
            'urllib3': {
                'level': 'INFO',
            },
        },
    }

    logging.config.dictConfig(logging_config)
    log = logging.getLogger(app_name)
    return log


if __name__ == "__main__":
    log = init_logging(__name__)
    main()
