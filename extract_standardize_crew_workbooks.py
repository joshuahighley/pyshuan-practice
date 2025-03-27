import pandas as pd
pd.set_option("future.no_silent_downcasting", True)
import numpy as np
import os
import difflib
import re

## For Google Colab workflow
# from google.colab import drive
# drive.mount('/content/drive/')
# folder = '/content/drive/MyDrive/Crew Workbooks for Analysis/'

import builtins
print = builtins.print
list = builtins.list

# Import the collection of crew workbook excel files
folder = 'Workbooks/'
crew_workbooks = os.listdir(folder)

print(crew_workbooks)






# Creates the basis for column names for renaming & organization down the line
# Columns in 'a' and 'j' type workbooks
erFile = 'Crew Notes - Job Costing SCHMIDT, RAEKELL - 184603.xlsm'
erFileDF = pd.read_excel(folder+erFile, 'Time Worksheet', header=None)
erHeader = erFileDF.iloc[2, 0:9]

# Columns in 'n' type workbooks, the latest layout/format for crew workbooks
newFile = 'New Crew Workbook - 957405 - 7-16-24_PY24.xlsm'
newFileDF = pd.read_excel(folder+newFile, 'Time Worksheet', header=None)
newHeader = newFileDF.iloc[6, 0:11]

# Indicate number of workbooks found in the folder
print(str(len(crew_workbooks)) + " workbooks found.")






# Use SequenceMatcher to compare strings; matching if similarity is above 0.8 (80% similar)
# Used to compare the headers of a given workbook to the two 'OG' headers (erHeader and newHeader)
def compare_ranges(range1, range2, threshold=5):
    matches = 0
    for str1, str2 in zip(range1, range2):

        similarity = difflib.SequenceMatcher(None, str(str1), str(str2)).ratio()
        if similarity > 0.7:
            matches += 1

    return matches >= threshold, matches

# Lists to sort workbooks and dataframes into
j_wbs = []
a_wbs = []
n_wbs = []

j_dfs = []
a_dfs = []
n_dfs = []

# Loop through each workbook in the folder, comparing the headers to the two 'OG' headers
# If a match is found, add the workbook to the appropriate list
for wb in crew_workbooks:
    try:
        df1 = pd.read_excel(folder+wb, 'Time Worksheet', header=None)
        a_are_similar, num_matches = compare_ranges(erHeader, df1.iloc[2, 0:9])
        j_are_similar, num_matches = compare_ranges(erHeader, df1.iloc[2, 9:18])
        n_are_similar, num_matches = compare_ranges(newHeader, df1.iloc[6, 0:11])
        if j_are_similar == True:
            j_wbs.append(wb)
            j_dfs.append(df1)
        elif a_are_similar == True:
            a_wbs.append(wb)
            a_dfs.append(df1)
        elif n_are_similar == True:
            n_wbs.append(wb)
            n_dfs.append(df1)
    except:
        None

# List how many of each "type" of workbooks are included
print(f"J: {str(len(j_wbs))}\nA: {str(len(a_wbs))}\nN: {str(len(n_wbs))}")






# Data Extractors for collecting crew data from each "type" of workbook
# J-Type (two-column), makes one dataframe from two columns
def j_extractor(j):
  col_names = j.iloc[2, 0:9]
  df1 = j.iloc[3:19, 0:9].reset_index(drop=True)
  df1.columns = col_names
  df2 = pd.DataFrame(j.iloc[3:19, 9:18]).reset_index(drop=True)
  df2.columns = col_names

  # Combined, unclean dataframe of extracted crew data
  df = pd.concat([df1, df2], axis=0).reset_index(drop=True)

  return df

# A-Type (one column)
def a_extractor(a):
  col_names = a.iloc[2, 0:9]
  df = a.iloc[3:107, 0:9].reset_index(drop=True)
  df.columns = col_names

  return df

# N-Type workbooks ("New Workbooks")
def n_extractor(n):
  col_names = n.iloc[6, 0:11]
  df = n.iloc[7:19, 0:11].reset_index(drop=True)
  df.columns = col_names


# Other functions for later:
# Find House IDs by regex patterns in workbook file names
hhn_list_by_pattern = []
hhn_errors = []

def hhn_checker(file_name):
    hhn_pattern = "\d{7}|\d{6}[abcd]|\d{6}|\d{5}"
    hhn = str(file_name).lower()
    check = re.search(hhn_pattern, hhn)
    if check != None:
        return re.findall(hhn_pattern, hhn)


# Clean a dataframe of crew data by removing NaNs, combining sublines below Measure #'s into one line
def clean_dataframe(df, columns_list):
    # Ensure 'Measure #' is the first column in the list, forward fill the 'Measure #' column
    if 'Measure #' in columns_list:
        columns_list.remove('Measure #')
        columns_list = ['Measure #'] + columns_list
    df['Measure #'] = df['Measure #'].ffill(downcast=None)

    # Step 2: Prepare the aggregation dictionary
    agg_dict = {}
    for col in columns_list[1:]:  # Skip 'Measure #'
        agg_dict[col] = lambda x: tuple(filter(lambda y: pd.notna(y) and y != '', x))

    # Step 3: Group by 'Measure #' and aggregate the data, clean up the columns, remove rows where all columns except 'Measure #' are NaN
    result = df.groupby('Measure #', as_index=False).agg(agg_dict)
    for col in columns_list[1:]:
        result[col] = result[col].apply(lambda x: x[0] if len(x) == 1 else x if len(x) > 1 else np.nan)
    result = result.dropna(subset=columns_list[1:4], how='all')

    return result


# Clean a dataframe of crew data by removing NaNs, combining sublines below Measure #'s into one line
def clean_j_dataframe(df, columns_list):
    # Ensure 'Measure #' is the first column in the list, forward fill the 'Measure #' column
    if 'Measure #' in columns_list:
        columns_list.remove('Measure #')
        columns_list = ['Measure #'] + columns_list
    df['Measure #'] = df['Measure #'].ffill(downcast=None)

    # Step 2: Prepare the aggregation dictionary
    agg_dict = {}
    for col in columns_list[1:]:  # Skip 'Measure #'
        agg_dict[col] = lambda x: tuple(filter(lambda y: pd.notna(y) and y != '' and y != 0.00, x))

    # Step 3: Group by 'Measure #' and aggregate the data, clean up the columns
    result = df.groupby('Measure #', as_index=False).agg(agg_dict)
    for col in columns_list[1:]:
        result[col] = result[col].apply(lambda x: x[0] if len(x) == 1 else x if len(x) > 1 else np.nan)

    # Step 4: Remove rows where any column (except 'Measure #') contains NaN or 0.00
    result = result.dropna(subset=columns_list[1:])
    result = result[~(result[columns_list[1:]] == 0.00).any(axis=1)]

    return result


# Display number of Crew Workbooks to be processed
j_count = len(j_wbs)
a_count = len(a_wbs)
n_count = len(n_wbs)

print(f'{str(j_count+a_count+n_count)} crew workbooks (J: {str(j_count)}, A: {str(a_count)}, N: {str(n_count)})')

check_hhns = []

for j in crew_workbooks:
    hhn = hhn_checker(j)
    check_hhns.append(hhn[0])

print(f'{abs(len(set(check_hhns))-len(crew_workbooks))} Duplicate House IDs found out of {len(crew_workbooks)}.')

# Variables/Lists for Later
clean_dfs = []
hhn_check_list = []
empty_dfs = []






# Clean extracted crew data for J-Type crew workbooks
wb_number = 0
j_clean_dfs = []
j_hhns = []

print(len(j_wbs))

for df in j_dfs:
  df_raw = j_extractor(df) # function to extract
  cols = df_raw.columns.tolist() # collects column names for cleaning function
  df_clean = clean_dataframe(df_raw, cols) # function to clean data

  # Set House ID column to = what's in cell A2 (hopefully the HHN)
  hhn = hhn_checker(str(j_wbs[wb_number]))
  df_clean['House ID'] = str(hhn[0])

  # Sets Indexes to HHN and Measure #, respectively
  df_clean_indexed = df_clean.set_index(['House ID', 'Measure #'])
  df_clean_indexed.drop(df_clean_indexed.columns[4], axis=1, inplace=True)

  if df_clean.size > 0:
    j_clean_dfs.append(df_clean_indexed)
    clean_dfs.append(df_clean_indexed)
    j_hhns.append(hhn[0])
  else:
    empty_dfs.append(df_clean_indexed)

  wb_number += 1


print(len(j_clean_dfs))
print(j_clean_dfs)
print(len(set(j_hhns)))






# Clean extracted crew data for A-Type crew workbooks
wb_number = 0
a_clean_dfs = []
a_hhns = []
print('A-type workbooks to process: ' + str(len(a_dfs)))

print(a_wbs)

for df in a_dfs:
  df_raw = a_extractor(df) # function to extract
  cols = df_raw.columns.tolist() # collects column names for cleaning function
  df_clean = clean_dataframe(df_raw, cols) # function to clean data

  # Set House ID column to ...
  hhn = hhn_checker(a_wbs[wb_number])
  df_clean['House ID'] = hhn[0]

  # Sets Indexes to HHN and Measure #, respectively
  df_clean.drop(df_clean.columns[5], axis=1, inplace=True)
  df_clean_indexed = df_clean.set_index(['House ID', 'Measure #'])


  if df_clean.size > 0:
    clean_dfs.append(df_clean_indexed)
    a_clean_dfs.append(df_clean_indexed)
    a_hhns.append(hhn[0])
  else:
    empty_dfs.append(df_clean_indexed)

  wb_number += 1


print('A-type workbooks processed: ' + str(wb_number))
print(len(set(a_hhns)))
print(len(clean_dfs))
print(a_clean_dfs)






## N-Type Workbooks ("New Workbooks")
problem_wbs = []
n_clean_dfs = []
n_hhns = []
wb_number = 0


for n in n_dfs:
  mc = n.loc[:,0:0]
  mc.columns = {'num'}
  hhn = hhn_checker(str(n_wbs[wb_number]))

  try:
    m1_index = mc.loc[mc['num'] == 'Program Support'].index.values.astype(int)[0]
    m25_index = mc.loc[mc['num'] == 'Do Not Touch'].index.values.astype(int)[0]
  except:
    measure1_index = 11
    measure25_index = 157

  n_df = n.iloc[m1_index:m25_index-1, 0:13].reset_index(drop=True)
  n_df.columns = n.iloc[6, 0:13]
  n_df.drop(n_df.columns[[1,6,10]], axis=1, inplace=True)

  n_df_clean = clean_j_dataframe(n_df, n_df.columns.tolist())
  n_df_clean['House ID'] = str(hhn[0])
  n_df_indexed = n_df_clean.set_index(['House ID', 'Measure #'])

  if n_df_indexed.size > 0:
    n_clean_dfs.append(n_df_indexed)
    clean_dfs.append(n_df_indexed)
    n_hhns.append(hhn[0])
  else:
    empty_dfs.append(n_df_indexed)
  
  wb_number += 1



print(n_clean_dfs)
print(len(n_clean_dfs))







# Collects a measure's materials and quantities from its workbook's measure sheet (e.g. 'M3')
# Intended to be applied to the clean dataframes, create a 'cell' for each row with list of materials, and one for number used
def measure_materials(row):
    hhn = row.name[0]       # House ID
    mn = row.name[1]        # Measure Number
    file = None
    
    # Determine the correct workbook file based on the House ID
    for wb in crew_workbooks:
        if wb.find(hhn) != -1 or wb.find(hhn.upper()) != -1:
            file = wb
            break
    
    print(f'Processing {hhn} - {mn} from {file}')
    
    try:
        df = pd.read_excel(f'Workbooks/{file}', f'm{str(mn)}', header=None)
    except ValueError as ve:
        try:
            df = pd.read_excel(f'Workbooks/{file}', f'M{str(mn)}', header=None)
        except Exception as e:
            return None, None
    except Exception as e:
        return None, None
    
    materials_col = df.iloc[2:20, 4]
    materials_clean = materials_col[materials_col != '-'].dropna().values
    quantity_col = df.iloc[2:20, 3].dropna()
    quantity_clean = quantity_col.values

    return materials_clean, quantity_clean



# Apply the measure_materials function to each row of the clean dataframes
materials_dfs = []

for i, wb in enumerate(clean_dfs):
    check = 0
    if check == 1:
        try:
            measure_data = wb.apply(measure_materials, axis=1, result_type='expand')
            measure_data.columns = ['Materials', 'Quantity']
            wb = pd.DataFrame.merge(wb, measure_data, left_index=True, right_index=True)
            materials_dfs.append(wb)
        except FileNotFoundError:
            print(hhn_checker(wb.name[0]))
    else:
        print('Skipping measure data extraction for this workbook.')
    
    print(f'\n{i+1}/{len(clean_dfs)}\n')




test = materials_dfs[30]

print(test)

for df in materials_dfs:
    hhn = (df.index.get_level_values('House ID').unique().values[0]).lower()
    directory = 'Output 1 - Extract/Pickles with Materials/'
    df.to_pickle(directory + hhn + '.pkl')






workbook_dfs = materials_dfs
check = 0
check_list = []

# Iterate through all workbooks to create a list of column names from Time Worksheets
for wb in workbook_dfs:
    wb_columns = wb.columns.tolist()
    for c in wb_columns:
        check_list.append(c)

# Create a list of unique column names
col_list = list(set(check_list))


nan_val = 0
col_check = 0

# Find the/a column with a value of NaN, record its index for later use
for n in col_list:
    if type(n) is str:
        col_check += 1
    else:
        nan_val = col_check
        col_check += 1






## COLUMN NAME STANDARDIZER
# Map with column names and their permutations to later rename workbook's dataframes with consistent column names
column_mapping = {
    'Measure Name': ['Measure Name', 'NOTES'],
    'Measure Type': ['Type', 'Measure Type', 'Measure', 'Measure Type - Use Drop Down'],
    'Estimated Hours': ['Estimated Hrs', 'Hrs per measure'],
    'Actual Hours': ['Actual Hrs', 'hours'],
    'Notes': ['Deductions', col_list[nan_val],]
}

# Create inverted map for easy lookup
inverted_column_map = {old: new for new, old_list in column_mapping.items() for old in old_list}

# The function that standardizes the column names for the workbook dataframes using the inverted column map
def standardize_columns(df):
    df_standard = df.rename(columns=inverted_column_map, inplace=False)
    for col in column_mapping.keys():
        if col not in df_standard.columns:
            df_standard[col] = pd.NA
    if 'Actual Hrs' not in df.columns and 'hours' not in df.columns:
        df_standard['Actual Hours'] = df_standard['Crew Hours'].apply(sum_hours)
    if 'Color Identifier' in df_standard.columns:
        df_standard.drop(columns=['Color Identifier'], inplace=True)
    
    df_standard['Crew'] = df_standard['Crew'].fillna('', inplace=False)
    df_standard['Crew'] = df_standard['Crew'].apply(lambda x: list(x) if type(x) == tuple else [x])
    df_standard['Crew'] = df_standard['Crew'].apply(correct_names)
    return df_standard






## CREW NAME STANDARDIZER
# Map with crew names and common misspellings to later clean the crew values in workbook dataframes
name_mapping = {
    'Chris': ['Chris'],
    'Emmett': ['EMMETT', 'Emmett', 'Emmettt'],
    'Erich': ['Erich'],
    'Jerry': ['JErry', 'Jerry'],
    'Julian': ['Julian'],
    'Lindsey': ['LIndsey', 'Lindsey'],
    'Leo': ['Leo'],
    'Mark': ['MArk', 'Mark', 'Mark ', 'mark'],
    'Mike': ['MIKE', 'Mike'],
    'Mao': ['Mao'],
    'Mohan': ['Mohan', 'mohan'],
    'Robert': ['Robert'],
    'Sam': ['Sam'],
    'Shuntilla': ['Shuntilla'],
    'Travis': ['TRAVIS', 'Traivs', 'Travis'],
    'Vue': ['Vue'],
    'Vuepao': ['Vuepao'],
    'Will': ['WIll', 'Will', 'Will ', 'will'],
    'Zach': ['Zach'],
    'Zachary': ['Zachary']
}

# Invert crew name map for easy lookup
inverted_name_map = {new: old for old, new_list in name_mapping.items() for new in new_list}

# Function to correct crew names with the inverted name map
def correct_names(name_list):
    return [inverted_name_map.get(name, name) for name in name_list]



## HOURS CLEANER
# Sum the crews' hours for a given measure, which are contained in tuples due to excel workbooks often having multiple lines per measure
def sum_hours(tup):
    if type(tup) == tuple:
        return sum(tup)
    else:
        return tup






# Standardize workbook columns
workbook_dfs_standard = []

for df in workbook_dfs:
    workbook_dfs_standard.append(standardize_columns(df))

check_num = 0

print('After Column Standardizer Function:')
print(workbook_dfs_standard)


# Create a list of all unique crew names
check = 0
check_list = ['Will', 'Leo']

for df in workbook_dfs_standard:
    df_crews = df['Crew'].dropna()
    #print(check)
    #print(df_crews)
    for x in df_crews:
        if type(x) is list:
            for y in x:
                if y not in check_list:
                    check_list.append(y)
        else:
            check_list.append(x)
    check += 1

crew_list = sorted(list(set(check_list)))






# What is this?
# Is this just another creation of a list of unique column names??
# It does it after the standardization function (above), so that's kind of nice but ideally not needed
# It's a reduction from 18 to 10.

all_crew = []
all_columns = []

for df in workbook_dfs_standard:
    crew_list = df['Crew']
    for l in crew_list:
        for crew in l:
           all_crew.append(crew)

    column_list = df.columns.tolist()
    for column in column_list:
        all_columns.append(column)


print(len(set(all_columns)))
print(set(all_columns))

for df in workbook_dfs_standard:
    if all_columns[9] in df.columns:
        print(df[all_columns[9]])






## CLEANING MEASURE NAMES

# Common typos in measure names, followed by inverted typo map for easy lookup
# example: 'correct': ['incorrect1', 'incorrect2', 'incorrect3']
typos = {
    'additional': ['addt'],
    'address': ['adress'],
    'air seal': ['airseal', 'airsealing'],
    'attic': ['attich', 'atic'],
    'boundary': ['boundry', 'boudary'],
    'centerpoint': ['cpt', 'cntpt', 'cntrpt', 'ctrpt'],
    'combustion': ['cuombustion', 'combustion', 'combustion'],
    'common wall': ['commonwall'],
    'decommissioned': ['decommissioned', 'decommission'],
    'dense pack': ['densepack', 'denseapack'],
    'detector': ['dector', 'det'],
    'diagnostic': ['diagonstic', 'diagnotic'],
    'diagnostics': ['dianostic', 'diagnotics'],
    'diagnostics': ['dianostics'],
    'exhaust fan': ['exhaustfan'],
    'infiltration': ['infil', 'inf.', 'infiltraion', 'inflitration', 'infultration'],
    'insulation': ['insulation', 'insulaton', 'ins', 'insul', 'ins.', 'insulatio'],
    'knee wall': ['kneewall'],
    'peak': ['peek'],
    'reduction': ['red.', 'reducton', 'redctn', 'red', 'redcn', 'reduc', 'reductio', 'red', 'reducton'],
    'replacement': ['replacment'],
    'retrofits': ['redrofits'],
    'sill box': ['sillbox'],
    'trampled': ['traampled', 'trmpled'],
    'wrap': ['wrao', 'wrp', 'wrape']
}

inverted_typo_map = {old: new for new, old_list in typos.items() for old in old_list} 

# function that takes input string, makes lower case w/ unusual spaces and typos removed
def clean_string(s):
    if isinstance(s, str):
        # Convert to lowercase
        s = s.lower()
        # Remove trailing periods (e.g. 'Ins.' becomes 'Ins') NOT WORKING AS INTENDED?
        s = re.sub(r'(?<!\d)\.(?!\d)', '', s)
        # Correct typos
        for typo, correction in inverted_typo_map.items():
            s = re.sub(r'\b(' + re.escape(typo) + r')\b', correction, s, flags=re.IGNORECASE)
        # Replace multiple spaces with a single space
        s = re.sub(r'\s+', ' ', s)
        # R-Value reformat
        s_1 = re.search(r'r[0-5][0-9]', s)
        if isinstance(s_1, re.Match):
            s_span = s_1.span()
            s = s[:(s_span[0]+1)] + '-' + s[(s_span[0]+1):]
        s_2 = re.search(r"r\s[0-9][0-9]", s)
        if isinstance(s_2, re.Match):
            s_span = s_2.span()
            s = s[:(s_span[0]+1)] + '-' + s[(s_span[0]+2):]
        # Strip leading and trailing whitespace
        return s.strip()
    else:
        return s


# Uses the clean_string function to clean the measure names in the workbook dataframes
for df in workbook_dfs_standard:
    df['Measure Name Unmodified'] = df['Measure Name']
    df['Measure Name'] = df['Measure Name'].apply(clean_string)






#print(n_clean_dfs)
print(len(workbook_dfs_standard))
pickle_count = 0
pickled_hhns = []

dfs_to_save = workbook_dfs_standard


for df in dfs_to_save[:]:
    hhn = (df.index.get_level_values('House ID').unique().values[0]).lower()
    print(hhn)
    directory = 'Output 1 - Extract/Pickles/'
    directory_csv = 'Output 1 - Extract/CSVs/'

    df.to_pickle(directory + hhn + '.pkl')
    df.to_csv(directory_csv + hhn + '.csv', index=False)


    #print(hhn)
    pickled_hhns.append(hhn)
    pickle_count += 1


# Combine workbook dataframes into one large dataframe, save as a CSV
wbs_unindexed = []

for df in dfs_to_save:
    df2 = df.reset_index(inplace=False)
    wbs_unindexed.append(df2)



base_df = pd.concat(wbs_unindexed, axis=0, ignore_index=True)

base_df.to_csv('Output 1 - Extract/All_Workbooks.csv', index=False)


print(base_df)
