# Usage: gcproc.py /path/to/data/directory "experiment name"
# 
# Author: Benjamin Chi
# Date: 2022-07-03
# Purpose: process multiple GC-FID data files and output the peak areas per analyte to an excel workbook for analysis.
# Note: an excel file named "cf.xls" must be kept in the same folder as the data "working" directory. This should contain a list of analytes with
# Front and Back correction factors followed by Front and Back retention times on adjacent columns for a total of 5 columns. Folders "*.D" should be
# in the same folder as well and contain the "Report.TXT" file.

import subprocess
import xlsxwriter
import xlrd
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
    # Append peak reference 'sample' to start of analyte tables
    all_analyte_tables.insert(0, peak_reference)

    # For multiple samples with different number of peaks, find the maximum number of peaks
    max_peak_num = len(all_analyte_tables[0][2])
    print("max_peak_num = %d" % max_peak_num)
    for i in range(1, len(all_analyte_tables)):
        peak_num = len(all_analyte_tables[i][2])
        if (max_peak_num < peak_num):
            max_peak_num = peak_num
            print("new max_peak_num = %d" % max_peak_num)
    
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

    for row in range(1, rows): # Account for empty cell before rows begin
        correction_factors.append([worksheet.cell_value(row, 0), worksheet.cell_value(row, 3), worksheet.cell_value(row, 4)])
    
    return correction_factors

# Returns array of form ['analyte name', 'front_ret', 'back_ret']
def get_ret_times(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    ret_times = []

    for row in range(1, rows): # Account for empty cell before rows begin
        ret_times.append([worksheet.cell_value(row, 0), worksheet.cell_value(row, 1), worksheet.cell_value(row, 2)])
    
    return ret_times

# Return name list
def get_names(cf_file):
    workbook = xlrd.open_workbook(cf_file)
    worksheet = workbook.sheet_by_index(0)

    rows = worksheet.nrows
    
    names = []

    for row in range(1, rows): # Account for empty cell before rows begin
        names.append([worksheet.cell_value(row, 0)])
    
    return names 

def write_xl(working_dir, experiment_name, data, cf, analytes):
    workbook = xlsxwriter.Workbook(working_dir + "/" + experiment_name + '_yields.xlsx')
    worksheet = workbook.add_worksheet()

    # Create formats
    header = workbook.add_format({'bold': True, 'underline': True})
    bold = workbook.add_format({'bold': True})
    percent = workbook.add_format({'num_format': '0.00%'})
    
    r_cf_start = 0
    c_cf_start = 0

    # Write in correction factors
    worksheet.write(r_cf_start, c_cf_start, "Correction Factors", header)
    worksheet.write(r_cf_start + 1, c_cf_start, "Reagent", bold)
    worksheet.write(r_cf_start + 1, c_cf_start + 1, "Front", bold)
    worksheet.write(r_cf_start + 1, c_cf_start + 2, "Back", bold)

    print(cf)
    for entry in range(0, len(cf)):
        for i in range(0, 3):
            worksheet.write(entry + r_cf_start + 2, c_cf_start + i, cf[entry][i])

    # Write in area data and headers
    r_data_start = 0
    c_data_start = 6
    
    r_is = 0
    c_is = 0
    
    worksheet.write(r_data_start, c_data_start, "GC Corrected Yields", header)
    worksheet.write(r_data_start + 1, c_data_start, "Detector", bold)
    worksheet.write(r_data_start + 1, c_data_start + 1, "Entry", bold)
    
    for analyte_col in range(0, len(analytes)): # write header names
    	name = analytes[analyte_col][0]
    	worksheet.write(r_data_start + 1, c_data_start + analyte_col + 2, name, bold)
    	if (re.match(".*_IS", name)):
    		r_is = r_data_start
    		c_is = c_data_start + analyte_col + 2
    		print("Found internal standard %s at: %s" % (name, get_cell(r_is, c_is)))
    		
    worksheet.set_column(c_data_start + 1, c_data_start + 1, 25)
    		
    for entry in range(0, len(data)): # write data
        columns = len(data[entry])
        for i in range(0, columns):
            cell = data[entry][i]
            try:
            	cell = float(cell)
            except:
            	pass
            worksheet.write(r_data_start + entry + 2, c_data_start + i, cell)
    
    # Write in formatted data
    r_form_start = r_data_start + len(data) + 4
    c_form_start = c_data_start + 2
    
    for entry in range(0, len(data)):
    	worksheet.write_formula(r_form_start + entry, c_form_start - 1, '=' + get_cell(r_data_start + 2 + entry, c_data_start + 1))
    	
    	columns = len(data[entry])
    	
    	for i in range(0, columns - 3):
            worksheet.write_formula(r_form_start + entry, c_form_start + i, get_formula(r_data_start + entry + 2, c_data_start, r_cf_start + i + 2, c_cf_start, c_is, i), percent)
            
    # Add conditional formatting
    for analyte_col in range(0, len(analytes)):
    	worksheet.conditional_format(get_cell(r_form_start, c_form_start + analyte_col) + ":" + get_cell(r_form_start + len(data), c_form_start + analyte_col),
    	{'type': 'data_bar', 'bar_solid': True, 'min_type': 'num', 'max_type': 'num', 'min_value': 0, 'max_value': 1.0})
    workbook.close()

# convert coordinate grid e.g. (0,0) to excel cell format e.g. "A1"
def get_cell(r, c):
    return str(chr(c + 65)) + str(r+1) # Account for cell row in xlsxwriter starts at 0

# r = first row of given data block, c = first column of given data block, e.g. "Front" or "Back".
def get_formula(r, c, r_cf, c_cf, c_is, column):
    if_num = '=(' + get_cell(r, c + 2 + column) + '/' + get_cell(r, c_is) + ')/'
    if_det = 'IF(' + get_cell(r, c) + '="Front", ' + get_cell(r_cf, c_cf + 1) + ',' + get_cell(r_cf, c_cf + 2) + ')'
    return if_num + if_det

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
    
    working_dir = sys.argv[1]
    experiment_name = sys.argv[2]
    
    ret_times = format_ret(working_dir + '/cf.xls')
    peak_reference_front = ["peaks", "Front", ret_times[0]]
    peak_reference_back = ["peaks", "Back", ret_times[1]]
    
    # extract reports and organize as back or front detector
    data_list = os.listdir(working_dir)
    data_folders = list(filter(re.compile(".*\.D").match, data_list))

    report_extracted_front = []
    report_extracted_back = []

    for i in range(len(data_folders)):
        extract = extract_report_txt(working_dir + "/" + data_folders[i] + "/Report.TXT")
        if (extract[1] == "Front"):
            report_extracted_front.append(extract)
        else:
            report_extracted_back.append(extract)

    # Generate input files for GCalignR and process data
    generate_input_file(sort_index(report_extracted_front, 0), working_dir + '/input_data_front.txt', peak_reference_front) 
    generate_input_file(sort_index(report_extracted_back, 0), working_dir + '/input_data_back.txt', peak_reference_back)
    front_areas = get_area(working_dir + '/input_data_front.txt')
    back_areas = get_area(working_dir + '/input_data_back.txt')

    # Add table headers and sort
    for entry in front_areas:
        entry.insert(0, "Front")
    for entry in back_areas:
        entry.insert(0, "Back")
    all_areas = sort_index(front_areas + back_areas, 1)
    
    # write data to excel workbook
    cf = get_corr_factors(working_dir + '/cf.xls')
    analytes = get_names(working_dir + '/cf.xls')
    write_xl(working_dir, experiment_name, all_areas, cf, analytes)


if __name__ == "__main__":
    main()
