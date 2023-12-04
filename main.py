# -------------------------------------------------------------------------------
# Name:        process_survey123_datafile
# Purpose:     post processing of Survey123 fish survey field data into a flat file format.
#
# Author:      Adrian Kitchingman.
# Updated By:  Alexander Hopper
#
# Created:     13/05/2022
# Copyright:   (c) ak34 2022
# Licence:     <your licence>
# Last Update: 04/12/2023
# -------------------------------------------------------------------------------


#==========================================================================================================================================#
#==========================================================Preference Changes:=============================================================#
#==========================================================================================================================================#

#The following is to order how each page is presented in the results. Change export_default to True to output ALL data.
export_default = False
#The data column will not be output if the index is == -1.
#The data column will move to the corresponding index value, e.g. if 'site_code' is at index 1,
#then [-1, 0, -1, -1, ...] will only output 'site_code' and place it at the start.
#A 'j' will join that column and the following column, and place it in the position of the value after the 'j'.
#A list of [0, 1, 2, 3, ...] will not change any order.
#This will also not take into account 'ObjectID', i.e. 'GlobalID' is the first element.
if export_default == True:
    survey_template = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]
    location_template = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] #Keep in mind [... x, y] will become ... x_start, y_start, x_end, y_end]
    shot_template = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]
    obs_template = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
    sample_template = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]

else:
    survey_template = [-1, 1, 4, 5, 'j', 6, 7, 8, 9, 10, 11, 12, 0, 13, -1, 14, 15, 16, 17, 18, 19, 20, 21, -1, -1, -1, -1, -1, 2, 3]
    location_template = [-1, -1, -1, -1, -1, -1, -1, 0, 1, 2, 3] #Keep in mind [... x, y] will become ... x_start, y_start, x_end, y_end]
    shot_template = [-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, -1, -1, -1, -1, -1]
    obs_template = [-1, -1, -1, 0, 1, 2, -1, -1, -1, -1, -1, -1]
    sample_template = [-1, -1, -1, 0, 1, 2, 3, 4, 5, 6, 7, -1, 8, 9, 10, -1, -1, -1, -1, -1]

#Names of columns to sort for raw and tally sheets:
raw_sorters = ['Site_GlobalID', 'survey_date', 'section_number', 'species', 'observed', 'section_collected']
tally_sorters = ['Site_ID', 'Section_Number', 'Species']

#==========================================================================================================================================#
#========================================================End Preference Changes:===========================================================#
#==========================================================================================================================================#

# import copy
# import csv
# import local_vars as localvars
import io
import os
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import process_survey123_field_data_classes as cls
import process_survey123_field_data_functions as func
from tkinter import *
import tkinter.messagebox
from tkinter.filedialog import askopenfilename #as fd
import random
from inspect import currentframe, getframeinfo

root = Tk()
root.withdraw()
root.update()

dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)

# io_path = localvars.io_path
filename = askopenfilename(initialdir=dir_path, title="Open Survey123 XLSX File")
in_xlfile = os.path.basename(filename)
io_path = filename.replace(in_xlfile, '')

while True:   # repeat until the try statement succeeds
    try:
        workbook = load_workbook(filename)
        break

    except IOError:

        answer = tkinter.messagebox.askokcancel("Open File Error", "Could not open file! Please close Excel. Press OK to retry.")
root.destroy()

print(workbook.sheetnames)

# # ## READ SITE SURVEY DATA =======================================================================================

survey_list = func.read_in_excel_tab(workbook.worksheets[0])
survey_list_header = list(func.read_in_excel_tab_header(workbook.worksheets[0]))

# # ## READ location DATA =======================================================================================

loc_list = func.read_in_excel_tab(workbook['site_location_repeat_1'])
loc_list_header = list(func.read_in_excel_tab_header(workbook['site_location_repeat_1']))

# # ## READ SHOT DATA =======================================================================================

shot_list = func.read_in_excel_tab(workbook['shot_repeat_2'])
shot_list_header = list(func.read_in_excel_tab_header(workbook['shot_repeat_2']))

# # ## READ observed DATA =======================================================================================

obs_list = func.read_in_excel_tab(workbook['observed_fish_repeat_3'])
obs_list_header = list(func.read_in_excel_tab_header(workbook['observed_fish_repeat_3']))

# # ## READ sampled DATA =======================================================================================

sample_list = func.read_in_excel_tab(workbook['fish_sample_repeat_4'])
sample_list_header = list(func.read_in_excel_tab_header(workbook['fish_sample_repeat_4']))

#Sort the samples so any defined shots are at the top.
sample_list.sort(key=lambda x: 0 if x[sample_list_header.index('section_number_samp')] is None else int(x[sample_list_header.index('section_number_samp')]), reverse=True)
sample_list.sort(key=lambda x: x[sample_list_header.index('ParentGlobalID')])


#Change x,y to x_start,y_start,x_end,y_end in location header.
loc_list_header[loc_list_header.index('x')] = 'x_coordinate'
loc_list_header[loc_list_header.index('y')] = 'y_coordinate'
loc_list_header.append('finish_x_coordinate')
loc_list_header.append('finish_y_coordinate')

#Prepare lists for output - raw data & tally results:
print('Gathering Data...')
raw_data = []
tally_results = []

#Loop through site survey sheet
for svy in survey_list:
    survey_list_current = list(svy)

    #Change gear_type name and set section_condition to (UN)FISHABLE
    survey_list_current[survey_list_header.index('gear_type')] = func.gear_types[survey_list_current[survey_list_header.index('gear_type')]]
    if survey_list_current[survey_list_header.index('section_condition')].lower() == 'yes':
        survey_list_current[survey_list_header.index('section_condition')] = 'FISHABLE'
    else:
        survey_list_current[survey_list_header.index('section_condition')] = 'UNFISHABLE'

    creator = survey_list_current[survey_list_header.index('Creator')]

    #Loop through locations survey sheet - filtering for each survey.
    for lcs in filter(lambda x: x[loc_list_header.index('ParentGlobalID')] == survey_list_current[survey_list_header.index('GlobalID')] , loc_list):
        loc_list_current = list(lcs)

        #Only add data if starting coordinate to not double up on data.
        if loc_list_current[loc_list_header.index('point_location')] == 'site_start':

            #See if there is another location entry with same GlobalID for end coordinates.
            pair = False
            x_finish = 0
            y_finish = 0
            for same_location in filter(lambda x:x[loc_list_header.index('GlobalID')] == loc_list_current[loc_list_header.index('GlobalID')] ,loc_list):
                if same_location[loc_list_header.index('point_location')] == 'site_end':

                    pair = True
                    x_finish = same_location[loc_list_header.index('x_coordinate')]
                    y_finish = same_location[loc_list_header.index('y_coordinate')]
            
            loc_list_current.append(x_finish)
            loc_list_current.append(y_finish)
                

            #Loop through shots sheet - filtering for location.
            sts = list(filter(lambda x: x[shot_list_header.index('ParentGlobalID')] == loc_list_current[loc_list_header.index('ParentGlobalID')], shot_list))
            if len(sts) > 0:

                for shot_list_current in sts:
                    shot_list_current = list(shot_list_current)

                    section_index = shot_list_header.index('section_number')
                    if shot_list_current[section_index] is None and len(sts) == 1:
                        shot_list_current[section_index] = 1
                    

                    #Create filler for samples:
                    sample_list_current = [None] * len(sample_list_header)

                    #Loop through observations - filtering for shots. And removing shots with 'obs_ts' == None.
                    obs = list(filter(lambda x: x[obs_list_header.index('ParentGlobalID')] == shot_list_current[shot_list_header.index('GlobalID')] and x[obs_list_header.index('obs_ts')] is not None, obs_list))
                
                    if len(obs) > 0:
                        for obs_list_current in obs:
                            obs_list_current = list(obs_list_current)

                            s_custom = obs_list_header.index('species_obs_custom')
                            s_new = obs_list_header.index('species_new')
                            s_obs = obs_list_header.index('species_obs')

                            if survey_list_current[survey_list_header.index('section_condition')] == 'FISHABLE':

                                #Build object for output
                                #Find ID Indices:
                                ID_Indices = [survey_list_header.index('GlobalID'),
                                         loc_list_header.index('GlobalID'),
                                         shot_list_header.index('GlobalID'),
                                         obs_list_header.index('GlobalID'),
                                         sample_list_header.index('GlobalID'),]
                                raw_data.append(cls.resultObject(survey_list_current, 
                                                        loc_list_current, 
                                                        shot_list_current, 
                                                        obs_list_current, 
                                                        sample_list_current,
                                                        creator,
                                                        ID_Indices))

                            #If section_collected is None, set to 0
                            if obs_list_current[obs_list_header.index('section_collected')] is None:
                                obs_list_current[obs_list_header.index('section_collected')] = 0

                            #Determine species from columns custom, new and obs.
                            if obs_list_current[s_custom] is not None or obs_list_current[s_new] is not None or obs_list_current[s_obs] is not None:
                                if obs_list_current[s_custom] is None:
                                    species = obs_list_current[s_obs]
                                elif obs_list_current[s_custom] is not None:
                                    species = obs_list_current[s_custom]
                                else:
                                    species = obs_list_current[s_new]

                                #Enter correct species
                                obs_list_current[s_obs] = species

                                # Set collected and observed to 0 if None
                                collected = obs_list_current[obs_list_header.index('section_collected')]
                                observed = obs_list_current[obs_list_header.index('observed')]
                                
                                if collected is None:
                                    obs_list_current[obs_list_header.index('section_collected')] = 0
                                if observed is None:
                                    obs_list_current[obs_list_header.index('observed')] = 0

                                collected = obs_list_current[obs_list_header.index('section_collected')]
                                observed = obs_list_current[obs_list_header.index('observed')]
                                
                                if collected != 0 or observed != 0:
                                    
                                    #Build object for tally results:
                                    site_id = survey_list_current[survey_list_header.index('GlobalID')]
                                    section_number = shot_list_current[section_index]
                                    #Species already defined
                                    #Collected already defined
                                    #Observed already defined
                                    shot_id = shot_list_current[shot_list_header.index('GlobalID')]
                                    obs_id = obs_list_current[obs_list_header.index('GlobalID')]
                                    #Creator already defined

                                    #Send tally data for output
                                    tally_results.append([site_id, section_number, species, collected, observed, collected, shot_id, obs_id, creator])
                                #Define tally header, but only once
                                try:
                                    tally_header
                                except:
                                    tally_header = ['Site_ID', 'Section_Number', 'Species', 'Collected', 'Observed', 'Collected_Tally', 'shot_id', 'obs_id', 'Creator']

                    #If no observations match, still output if fishable               
                    else:
                        if survey_list_current[survey_list_header.index('section_condition')] == 'FISHABLE':
                            if obs_list_current[obs_list_header.index('section_collected')] is None:
                                obs_list_current[obs_list_header.index('section_collected')] = 0

                            #Build object for output
                            #Find ID Indices:
                            ID_Indices = [survey_list_header.index('GlobalID'),
                                         loc_list_header.index('GlobalID'),
                                         shot_list_header.index('GlobalID'),
                                         obs_list_header.index('GlobalID'),
                                         sample_list_header.index('GlobalID'),]
                            raw_data.append(cls.resultObject(survey_list_current, 
                                                        loc_list_current, 
                                                        shot_list_current, 
                                                        obs_list_current, 
                                                        sample_list_current,
                                                        creator,
                                                        ID_Indices))

                        
            #If no shots exist, add 1 shot if fishable or samples present and add site information.
            else:
                #Create filler for observations:
                obs_list_current = [None] * len(obs_list_header)
                obs_list_current[obs_list_header.index('section_collected')] = 0

                #Loop through samples - filtering for shots.
                for smp in filter(lambda x: x[sample_list_header.index('ParentGlobalID')] == survey_list_current[survey_list_header.index('GlobalID')] , sample_list):
                    sample_list_current = list(smp)

                    #Find correct species:
                    if sample_list_current[sample_list_header.index('species_samp')] is None:
                        sample_list_current[sample_list_header.index('species_samp')] = sample_list_current[sample_list_header.index('species_samp_custom')]

                    # Build object for output
                    #Find ID Indices:
                    ID_Indices = [survey_list_header.index('GlobalID'),
                                         loc_list_header.index('GlobalID'),
                                         shot_list_header.index('GlobalID'),
                                         obs_list_header.index('GlobalID'),
                                         sample_list_header.index('GlobalID'),]
                    raw_data.append(cls.resultObject(survey_list_current, 
                                                    loc_list_current, 
                                                    shot_list_current, 
                                                    obs_list_current, 
                                                    sample_list_current,
                                                    creator,
                                                    ID_Indices))           
#Create checklist for sample data, at least one value must be NOT 0 or None or will skip.
sample_checklist = ['section_number_samp', 
                    'fork_length',
                    'total_length',
                    'weight',
                    'collected',
                    'recapture',
                    'external_tag_no',
                    'pit',
                    'genetics_label',
                    'otoliths_label',
                    'fauna_notes']

#Create random order to find random shot.
out_list_length = len(raw_data)
sampling_list = list(range(out_list_length))

#Loop through samples.
for smp in sample_list:
    sample_list_current = list(smp)
    creator = sample_list_current[sample_list_header.index('Creator')]
    s_custom = sample_list_current[sample_list_header.index('species_samp_custom')]
    s_samp = sample_list_current[sample_list_header.index('species_samp')]

    #Fix up species name
    if s_custom is not None or s_samp is not None: 
        species = s_samp if s_custom is None else s_custom
        
        sample_list_current[sample_list_header.index('species_samp')] = species
        skip_samp = TRUE
        sample_comparison = 0

        #Compare with checklist to determine whether to skip or output.
        for dummy in sample_checklist:
            if sample_list_current[sample_list_header.index(dummy)] not in [0 , None]:
                skip_samp = FALSE
        
        if skip_samp == FALSE:
            #Find random shot to attribute sample to with same ParentGlobalID
            #And create new object for final results.
            PGID = sample_list_current[sample_list_header.index('ParentGlobalID')]
            random.shuffle(sampling_list)
            for i in sampling_list:
                
                if PGID == raw_data[i].locations[loc_list_header.index('ParentGlobalID')] and species == raw_data[i].observations[obs_list_header.index('species_obs')]:

                    #Fix collected number:
                    section_num = (0 if raw_data[i].shots[shot_list_header.index('section_number')] is None else raw_data[i].shots[shot_list_header.index('section_number')])
                    func.adjust_species_count(sample_list_current, raw_data, PGID, section_num, species,survey_list_header, obs_list_header, sample_list_header, loc_list_header, shot_list_header, tally_results, tally_header)                                    
                    
                    #Build obs list:
                    obs_list_current = [None] * len(obs_list_header)
                    if sample_list_current[sample_list_header.index('collected')] is None or sample_list_current[sample_list_header.index('collected')] == 0:
                        obs_list_current[obs_list_header.index('section_collected')] = 1
                    else:
                        obs_list_current[obs_list_header.index('section_collected')] = sample_list_current[sample_list_header.index('collected')]
                    if obs_list_current[obs_list_header.index('observed')] is None:
                        obs_list_current[obs_list_header.index('observed')] = 0
                    
                    #Find ID Indices:
                    ID_Indices = [survey_list_header.index('GlobalID'),
                                         loc_list_header.index('GlobalID'),
                                         shot_list_header.index('GlobalID'),
                                         obs_list_header.index('GlobalID'),
                                         sample_list_header.index('GlobalID'),]
                    raw_data.append(cls.resultObject(raw_data[i].surveys, 
                                                raw_data[i].locations, 
                                                raw_data[i].shots, 
                                                obs_list_current, 
                                                sample_list_current,
                                                creator,
                                                ID_Indices))  
                    break

#Correctly format and order entries as list at top of code.
for obj in raw_data:

    survey_list_header_edit = obj.order(obj.surveys,survey_template,survey_list_header,'survey')
    loc_list_header_edit = obj.order(obj.locations,location_template,loc_list_header,'location')
    shot_list_header_edit = obj.order(obj.shots,shot_template,shot_list_header,'shot')
    obs_list_header_edit = obj.order(obj.observations,obs_template,obs_list_header,'obs')
    sample_list_header_edit = obj.order(obj.samples,sample_template,sample_list_header,'sample')

#Collate data and append GlobalID List and creator:
print('Collating data...')
raw_data_header = survey_list_header_edit + loc_list_header_edit + shot_list_header_edit + obs_list_header_edit + sample_list_header_edit
raw_data_header += ['Survey_GlobalID', 'Site_GlobalID', 'Shot_GlobalID', 'Obs_GlobalID', 'Sample_GlobalID', 'Creator']

for i in raw_data:
    i.collate(raw_data_header)

raw_data_header.pop(raw_data_header.index('species_samp'))
raw_data_header[raw_data_header.index('species_obs')] = 'species'

print('Processing data into new format...')

#Open new worksheet and create Raw Data & Tally Results
wb = openpyxl.Workbook()
ws_write = wb.active
ws_write.title = "Raw Data"
ws2_write = wb.create_sheet("Tally Results", 1)

#Write header and all data.
#Raw Data:
row_count = 1
func.write_row(ws_write, row_count, 1, raw_data_header)
for result in raw_data:
    row_count += 1
    func.write_row(ws_write, row_count, 1, tuple(result.collation))
#Tally Data:
row_count = 1
func.write_row(ws2_write, row_count, 1, tally_header)
for result in tally_results:
    row_count += 1
    func.write_row(ws2_write, row_count, 1, result)

#Sort both pages of output by column name at top of code:
try:
    #Calculates indices of sort columns:
    raw_sort_indices = []
    tally_sort_indices = []

    for name in raw_sorters:
        raw_sort_indices.append(raw_data_header.index(name) + 1)
    for name in tally_sorters:
        tally_sort_indices.append(tally_header.index(name) + 1)
    func.sheet_sort_rows(ws_write, 2, 0, raw_sort_indices)
    func.sheet_sort_rows(ws2_write, 2, 0, tally_sort_indices)
except:
    print("ERROR Sorting Columns, Check Sort Names")

#Resize columns to match entry lengths:
for column_cells in ws_write.columns:
    length = max(len(str(cell.value)) for cell in column_cells) + 3
    ws_write.column_dimensions[column_cells[0].column_letter].width = length
for column_cells in ws2_write.columns:
    length = max(len(str(cell.value)) for cell in column_cells) + 3
    ws2_write.column_dimensions[column_cells[0].column_letter].width = length

#Add auto filter to output sheet:
ws_write.auto_filter.ref = ws_write.dimensions
ws2_write.auto_filter.ref = ws2_write.dimensions


#Determine current date and export.
try:
    func.set_col_date_style(ws_write, raw_data_header.index('survey_date'))
except:
    frameinfo = getframeinfo(currentframe())
    print('ERROR editing survey date, check line', frameinfo.lineno - 2, 'for correct column title.')

dnow = datetime.now()
fdt = dnow.strftime("%y") + dnow.strftime("%m") + dnow.strftime("%d")

out_xlfile = in_xlfile.replace('(', '').replace(')', '')
out_xlfile = out_xlfile.replace(".xlsx", "_FlatFormat_" + fdt + ".xlsx")

wb.save(io_path + out_xlfile)

print('\nFormatting complete. New Excel file is at:\n{0}\n\n'.format(io_path + out_xlfile))


