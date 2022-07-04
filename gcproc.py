# Usage: gcproc.py /path/to/data/directory "experiment name"
# 
# Author: Benjamin Chi
# Date: 2022-07-03
# Version: 1.1.0
# Purpose: process multiple GC-FID data files and output the peak areas per analyte to an excel workbook for analysis.
# Note: an excel file named "cf.xls" must be kept in the same folder as the data "working" directory. This should contain a list of analytes with
# Front and Back correction factors followed by Front and Back retention times on adjacent columns for a total of 5 columns. Folders "*.D" should be
# in the same folder as well and contain the "Report.TXT" file.

import subprocess
import xlsxwriter
import xlrd
import json
import sys
import re
import os

# Align peaks and return peak area table
def get_area(input_file):
    command = '/usr/local/bin/Rscript'
    script = '/home/bkchi/Documents/gcproc/gcproc.R'
    args = input_file
    
    output = subprocess.check_output([command, script, args], universal_newlines=True)
    
    # Locate the start and end of the dataframe with analyte peak areas
    pattern = 'START(?:\s|.)*END' #?: non capaturing group so that re does not return matched capture group.
    result = re.findall(pattern, output)
    data = re.split('\n', result[0])
    data_trim = data[2 : len(data)-1] # remove Start and End headers
    
    # Remove white spaces and add data to table
    area_table = []

    for i in range(0, len(data_trim)):
        line = re.split('\s+', data_trim[i])
        area_table.append(line)
    
    return area_table

# Read in agilent REPORT.txt file and return array with the sample name as the first element,
# the detector as second element, and an array of format ['peak', 'area'] as the second element. Encoding is utf-16
def extract_report_txt(report_file):
    f = open(report_file, "r", encoding='utf-16')
    content = f.read()
    
    # Get Front or Back detector
    pattern = "\S+ Signal"
    detector = re.split('\s+', re.findall(pattern, content)[0])[0]

    # Get sample name
    pattern = 'Sample Name: .*'
    sample_name = re.split('\s+', re.findall(pattern, content)[0])[2]

    # Find the general pattern "Peak RetTime Sig Type Area"
    pattern = '\d+[ \t]+[\d\S]+[ \t]+\d+[ \t]+..[ \t]+[\d\S]+'
    result = re.findall(pattern, content)
    
    # Take only the RetTime and Area
    peak_num = len(result)
    analyte_table = []
    for i in range(0, peak_num):
        line = re.split('\s+', result[i])
        analyte_table.append([line[1], line[4]])
    
    return [sample_name, detector, analyte_table]

# Generate the tab-delimited input txt file for GCalignR using an array of analyte tables of type ['sample name',
# ['peak', 'area']], and output file path, and peak reference list to align to. The first two lines will be the sample
# names and column names ('RT', 'Area') respectively. The following lines will contain for each peak the peak-area pair
# for every sample separated by a tab.
def generate_input_file(all_analyte_tables, output_file, peak_reference):
    print("Generating input file for GCalignR at '%s'..." % output_file)
    
    # Append peak reference 'sample' to start of analyte tables
    all_analyte_tables.insert(0, peak_reference)

    # For multiple samples with different number of peaks, find the maximum number of peaks
    print("\nFinding maximum peak number across samples...")
    max_peak_num = len(all_analyte_tables[0][2])
    print("max_peak_num = %d" % max_peak_num)
    for i in range(1, len(all_analyte_tables)):
        peak_num = len(all_analyte_tables[i][2])
        if (max_peak_num < peak_num):
            max_peak_num = peak_num
            print("new max_peak_num = %d" % max_peak_num)
    print("Done.\n")
    
    # Initialize the input file information.
    sample_num = len(all_analyte_tables)
    sample_name = ""
    data = ""
    
    # For each peak, extract the peak-area pair for that peak in every sample and concat into one line, tab-delim.
    # Use try, catch block to add in tabs for samples with fewer than the maximum peak number.
    for i in range(0, max_peak_num):
        for j in range(0, sample_num):
            try:
                data += str(all_analyte_tables[j][2][i][0]) + "\t" + str(all_analyte_tables[j][2][i][1])
            except:
                data += "\t"
            if (j < sample_num - 1):
                data += "\t"
            else:
                data += "\n"
    
    # Collect sample names and format them as a single string, tab-delim.
    for i in range(0, sample_num):
        sample_name += re.sub("[-\.]", "_", all_analyte_tables[i][0])
        if (i < sample_num - 1):
            sample_name += "\t"
        else:
            sample_name += "\n"

    # Open and write the date to file.
    f = open(output_file, "w")
    f.write(sample_name)
    f.write("RT\tArea\n") # "RT" and "Area" are the column headers used in the gcproc.R script
    f.write(data)

# Sort two-dimensional array by index. Returns sorted array
def sort_index(report, index):
    convert = lambda text: int(text) if text.isdigit() else text
    alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key[index]) ]
    return sorted(report, key = alphanum_key)

# Returns array of form ['analyte name', 'front_cf', 'back_cf']
def get_corr_factors(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    correction_factors = []

    for row in range(2, rows): # Account for empty cells and headers before rows begin
        correction_factors.append([worksheet.cell_value(row, 0), worksheet.cell_value(row, 3), 
            worksheet.cell_value(row, 4), worksheet.cell_value(row, 6)])
    
    return correction_factors

# Returns array of form ['analyte name', 'front_ret', 'back_ret']
def get_ret_times(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    ret_times = []

    for row in range(2, rows): # Account for empty cells and headers before rows begin
        ret_times.append([worksheet.cell_value(row, 0), worksheet.cell_value(row, 1), 
            worksheet.cell_value(row, 2)])
    
    return ret_times

# Return name list
def get_names(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    names = []

    for row in range(2, rows): # Account for empty cells and headers before rows begin
        names.append(worksheet.cell_value(row, 0))
        
    return names

# Return color list as Hex code
def get_colors(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    colors = []

    for row in range(2, rows): # Account for empty cells and headers before rows begin
        colors.append(worksheet.cell_value(row, 6))
        
    return colors
    
# Return internal standard MW
def get_is_mw(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    return float(worksheet.cell_value(rows - 1, 5))

# convert coordinate grid e.g. (0,0) to excel cell format e.g. "A1"
def get_cell(r, c):
    return str(chr(c + 65)) + str(r+1) # Account for cell row in xlsxwriter starts at 0

# Gets formula for calculating corrected yields
# r = first row of given data block, c = first column of given data block, e.g. "Front" or "Back".
def get_formula(r, c, r_cf, c_cf, r_ratio, c_ratio, c_is, column):
    if_num = '=(' + get_cell(r, c + 2 + column) + '/' + get_cell(r, c_is) + ')/'
    if_det = 'IF(' + get_cell(r, c) + '="Front", ' + get_cell(r_cf, c_cf + 1) + ',' + get_cell(r_cf, c_cf + 2) + ')'
    apply_ratio = '*' + get_cell(r_ratio + 2, c_ratio + 5)
    return if_num + if_det + apply_ratio

def write_block(workbook, worksheet, row, col, title, header_list, data, fmt = None):
    # Create formats
    header = workbook.add_format({'font_name': 'Arial', 'bold': True, 'underline': True})
    bold = workbook.add_format({'font_name': 'Arial', 'bold': True, 'border': 1})
    normal = workbook.add_format({'font_name': 'Arial', 'border': 1})
    
    # Write in title and headers
    worksheet.write(row, col, title, header)
    header_col = 0
    for header in header_list:
    	worksheet.write(row + 1, col + header_col, header, bold)
    	header_col += 1
    
    # Write in data
    data_rows = len(data)
    for entry in range(0, data_rows):
    	for column in range(0, len(data[0])):
    	    if (fmt == None):
    	        fmt = normal
    	    cell = data[entry][column]
    	    try:
    	        cell = float(cell)
    	    except:
    	        pass
    	    worksheet.write(row + entry + 2, col + column, cell, fmt)
    
# Write data to excel workbook. Row and columns are defined as the top-left corner of the table including header and title.
# Data table comes in form [["Detector", "Notebook Code", area1, area2, ...], ...]
def write_xl(working_dir, experiment_name, data, cf, analytes, is_mw):
    workbook = xlsxwriter.Workbook(working_dir + "/" + experiment_name + '_yields.xlsx')
    worksheet = workbook.add_worksheet()

    # Write in correction factors
    r_cf_start = 0
    c_cf_start = 0
    
    cf_header_list = ["Reagent", "Front", "Back", "Hex Code"]
    cf_title = "Correction Factors"
    write_block(workbook, worksheet, r_cf_start, c_cf_start, cf_title, cf_header_list, cf)
    
    # Write in internal standard amounts per entry
    r_mass_start = r_cf_start + len(cf) + 4
    c_mass_start = 0
    
    # Build internal standard amounts mass formulas table
    is_mass_data = []
    for entry in range(0, len(data)):
        mass_cell = get_cell(r_mass_start + entry + 2, c_mass_start + 1)
        mw_cell = get_cell(r_mass_start + entry + 2, c_mass_start + 2)
        is_cell = get_cell(r_mass_start + entry + 2, c_mass_start + 3)
        rxn_cell = get_cell(r_mass_start + entry + 2, c_mass_start + 4)
        is_mass_data.append([data[entry][1], 0, is_mw, "=" + mass_cell + "/" + mw_cell, 0, "=" + is_cell + "/" + rxn_cell])
        
    mass_header_list = ["Notebook Code", "IS mass (mg)", "IS MW (g/mol)", "IS (mmol)", "Rxn (mmol)", "IS/Rxn"]
    mass_title = "Internal Standard (IS) Added"
    write_block(workbook, worksheet, r_mass_start, c_mass_start, mass_title, mass_header_list, is_mass_data)
    
    # Add number formatting to internal standard amounts table
    columns = len(is_mass_data[0])
    for col in range(0, columns - 1):
        start_cell = get_cell(r_mass_start + 2, c_mass_start + 1)
        end_cell = get_cell(r_mass_start + len(is_mass_data) + 1, c_mass_start + columns)
        
        print("Applying conditional formatting for range %s:%s..." % (start_cell, end_cell))
        
        num = workbook.add_format({'num_format': '0.00'})
        
        worksheet.conditional_format(start_cell + ":" + end_cell, {'type': 'no_errors',
                                                                   'format': num})
    
    # Write in area data and headers
    r_data_start = 0
    c_data_start = 7
    
    r_is = 0
    c_is = 0
    
    data_header_list = ["Detector", "Notebook Code"] + analytes
    data_title = "GC-FID Analyte Areas"
    write_block(workbook, worksheet, r_data_start, c_data_start, data_title, data_header_list, data)
    
    # Find the internal standard column and top row.
    for analyte_col in range(0, len(analytes)):
    	name = analytes[analyte_col]
    	if (re.match(".*_IS", name)):
    		r_is = r_data_start
    		c_is = c_data_start + analyte_col + 2
    		print("\nFound internal standard %s at: %s\n" % (name, get_cell(r_is, c_is)))
    
    # Write in formatted data
    r_form_start = r_data_start + len(data) + 4
    c_form_start = c_data_start
    
    form_title = "GC-FID Corrected Yields"
    form_header_list = ["Entry", "Notebook Code"] + analytes[:len(analytes)-1] + ["MB"] # Remove internal standard header and replace with mass balance.
    form_data = []
    
    # Make formatted data and exclude internal standard column - replace with mass balance.
    for entry in range(0, len(data)):
        row = ["Enter entry here."] # Initialize and build row array
        row += ['=' + get_cell(r_data_start + 2 + entry, c_data_start + 1)]
        
        analyte_cols = len(data[0]) - 3 # This excludes the internal standard row - must be last row
        
        for col in range(0, analyte_cols):
            row += [get_formula(r_data_start + entry + 2, c_data_start, r_cf_start + col + 2, c_cf_start, r_mass_start + entry, c_mass_start, c_is, col)]
        
        row += ["=SUM(" + get_cell(r_form_start + entry + 2, c_form_start + 2) + ":" + get_cell(r_form_start + entry + 2, c_form_start + analyte_cols + 1) + ")"] # Add on mass balance formula at the end
        
        form_data.append(row)
    
    write_block(workbook, worksheet, r_form_start, c_form_start, form_title, form_header_list, form_data)
            
    # Add conditional formatting to corrected yields
    for analyte_col in range(0, len(analytes)):
        start_cell = get_cell(r_form_start + 2, c_form_start + analyte_col + 2)
        end_cell = get_cell(r_form_start + len(data) + 1, c_form_start + analyte_col + 2)
        
        print("Applying conditional formatting for range %s:%s..." % (start_cell, end_cell))
        
        percent = workbook.add_format({'num_format': '0.00%'})
        color = '#' + cf[analyte_col][3]
        
        worksheet.conditional_format(start_cell + ":" + end_cell, {'type': 'no_errors',
                                                                   'format': percent})
        worksheet.conditional_format(start_cell + ":" + end_cell, {'type': 'data_bar', 
                                                                   'bar_solid': True,
                                                                   'min_type': 'num', 
                                                                   'max_type': 'num', 
                                                                   'min_value': 0, 
                                                                   'max_value': 1.0,
                                                                   'bar_color': color})
    
    # Set column widths
    worksheet.set_column(c_cf_start, c_cf_start, 25)
    worksheet.set_column(c_cf_start + 1, c_cf_start + 1, 12)
    worksheet.set_column(c_cf_start + 2, c_cf_start + 2, 12)
    worksheet.set_column(c_cf_start + 3, c_cf_start + 3, 12)
    worksheet.set_column(c_cf_start + 4, c_cf_start + 4, 12)
    worksheet.set_column(c_cf_start + 5, c_cf_start + 5, 12)
    
    worksheet.set_column(c_data_start, c_data_start, 25)
    worksheet.set_column(c_data_start + 1, c_data_start + 1, 25)
    	
    workbook.close()

# format retention time array for GCalignR    
def format_ret(cf_file):
    ret_times = get_ret_times(cf_file)
    
    front_ret = []
    back_ret = []
    
    for entry in ret_times:
    	front_ret.append([entry[1], 0])
    	back_ret.append([entry[2], 0])
    
    return [front_ret, back_ret]

def main():
    
    # Get arguments - working directory and experiment name
    args = len(sys.argv) - 1
    working_dir = ""
    experiment_name = ""
    cf_file = ""
    
    if (args == 0):
        working_dir = input("Enter working directory: ")
        cf_file = input("Enter correction factor file name: ")
        experiment_name = input("Enter experiment name/folder: ")
    elif (args == 1):
        experiment_name = sys.argv[1]
        try:
            working_dir = json.load(open(os.getcwd() + '/config.json'))['working_directory']
            print('Working directory not supplied. Loaded: "%s"' % working_dir)
            
            cf_file = json.load(open(os.getcwd() + '/config.json'))['cf_file_name']
        except:
            pass
        print('Found 1 argument: %s, using as experiment name/folder' % sys.argv[1])
    elif (args == 3):
        working_dir = sys.argv[1]
        cf_file = sys.argv[2]
        experiment_name = sys.argv[3]
        print('Found 3 arguments: "%s", "%s", "%s"' % (sys.argv[1], sys.argv[2], sys.argv[3]))
    else:
        sys.exit("Usage:\tgcproc.py working_path cf_file experiment_name\n\tgcproc.py experiment_name") 
    
    # Append experiment_name to working_dir to get data location. Define cd_dir to be in cf folder of working directory
    data_dir = working_dir + "/" + experiment_name
    cf_dir = working_dir + "/cf/" + cf_file
    
    # Get Front and Back retention times
    ret_times = format_ret(cf_dir)
    peak_reference_front = ["peaks", "Front", ret_times[0]]
    peak_reference_back = ["peaks", "Back", ret_times[1]]
    
    # extract reports and organize as back or front detector
    data_list = os.listdir(data_dir)
    data_folders = list(filter(re.compile(".*\.D").match, data_list))

    report_extracted_front = []
    report_extracted_back = []

    for i in range(len(data_folders)):
        extract = extract_report_txt(data_dir + "/" + data_folders[i] + "/Report.TXT")
        if (extract[1] == "Front"):
            report_extracted_front.append(extract)
        else:
            report_extracted_back.append(extract)

    # Generate input files for GCalignR and process data
    generate_input_file(sort_index(report_extracted_front, 0), data_dir + '/input_data_front.txt', peak_reference_front) 
    generate_input_file(sort_index(report_extracted_back, 0), data_dir + '/input_data_back.txt', peak_reference_back)
    front_areas = get_area(data_dir + '/input_data_front.txt')
    back_areas = get_area(data_dir + '/input_data_back.txt')

    # Add table headers and sort
    for entry in front_areas:
        entry.insert(0, "Front")
    for entry in back_areas:
        entry.insert(0, "Back")
    all_areas = sort_index(front_areas + back_areas, 1)
    
    # write data to excel workbook
    cf = get_corr_factors(cf_dir)
    analytes = get_names(cf_dir)
    is_mw = get_is_mw(cf_dir)
    write_xl(data_dir, experiment_name, all_areas, cf, analytes, is_mw)
    
    # Save working directory
    config = {'working_directory': working_dir, 'cf_file_name': cf_file}
    json.dump(config, open(os.getcwd() + '/config.json', 'w'))


if __name__ == "__main__":
    main()
