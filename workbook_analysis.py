import pandas as pd
import numpy as np
from os import listdir
list = __builtins__.list

input_folder = 'Output 2 - Unify'

# Load Workbook Pickles
workbooks = listdir(input_folder+'/Pickles')
print('Workbook Pickles Found: ' + str(len(workbooks)))


# Make list of dataframes from the pickles
workbook_dfs = []
for p in workbooks:
    df = pd.read_pickle(input_folder+'/Pickles/'+p)
    workbook_dfs.append(df)






# Aggregate every measure in the 260 workbooks, save as a CSV
every_measure = []

for df in workbook_dfs:
    measures = df['Measure Name'].to_list()
    for m in measures:
        every_measure.append(m)


wb_df_indexes = map(pd.DataFrame.reset_index, workbook_dfs)
full_measures_df = pd.concat(wb_df_indexes)

print(every_measure)
print(len(every_measure))
measure_df = pd.DataFrame(every_measure)
measure_df.to_csv('every_measure.csv', index=False)
full_measures_df.to_csv('full_measures.csv', index=False)






# Creates a dataframe with the name of every measure issued in the given crew workbooks
measure_names = []

# Removes blanks (which represent as Floats?)
for df in workbook_dfs:
    for measure in df.loc[:, 'Measure Name']:
        if type(measure) == float:
            continue
        else:
            measure_names.append(measure)

# Creates a Set of all measure names to filter out duplicates, writes that Set to a .csv file with all unique measure names
all_measures = pd.DataFrame(sorted(list(set(measure_names))))
all_measures.to_csv('all_measures.csv', index=False)

print(len(all_measures))






# Use Regex to filter 'rim' from words like 'trim' - r'\brim' where \b filters for 'rim' at the beginning of a word
# Type out the string OR use the list + the term's index (e.g. common_terms[1] = 'sill)
common_terms = ['rim', 'sill', 'infiltration', 'bypass', 'blow', 'cavity']
search_term = common_terms[5]
search_re = r'\b'+search_term

# terms to avoid in a querry for whatever reason. Same format as ^above^, avoid_terms[0] = silly term that will never filtered out DFs
avoid_terms = ['florgus', 'block', 'brick', 'duct', 'open', 'ceiling']
avoid_term = r'\b'+avoid_terms[5]

# Create a list of measures that match the search term above, where each item is comprised of the entire row from the measure's respective DF
matching_measures = []

for df in workbook_dfs:
    result = df[df['Measure Name'].str.contains(search_re, na=False)]
    result2 = result[~result['Measure Name'].str.contains(avoid_term, na=False)]
    matching_measures.append(result2)


# Filter empty dataframes from list
data = list(filter(lambda df: not df.empty, matching_measures))

# Concatenate all dataframes items in the list into one DF of measures matching search term
# What's the point of this? Does this do anything?
if data:
    base_df = pd.concat(data, axis=0, ignore_index=False)
else:
    base_df = pd.DataFrame()






# Import crew workbook excel files to scrape Measure data from M# sheets
folder = 'Workbooks/'
crew_workbooks = listdir(folder)

# Create an unindexed dataframe for easier access of HHN and Measure # line-by-line, and a copy of the working DF to preserve the data above
unindexed_df = base_df.reset_index()
base_with_materials_df = base_df

# Extract a list of tuples of (HHN, Measure #) from DF, then unzip into separate lists
measure_ids = unindexed_df.apply(lambda m: (m['House ID'], m['Measure #']), axis=1)
hhns, measure_nums = zip(*measure_ids)

# Find the Measure's sheet in the given excel file, extract the data
def measure_detail_scraper(wb, m, material):
    try:    # Read the workbook sheet associated with the Measure Number (case of 'm' varies by workbook, hence try/except)
        df = pd.read_excel(f'Workbooks/{str(wb)}', f'm{str(m)}', header=None)
    except:
        df = pd.read_excel(f'Workbooks/{str(wb)}', f'M{str(m)}', header=None)
    
    # Extract the quantity of MATERIAL specified (e.g. "CELLULOSE"), if no match, except: 'NOT FOUND'
    try:    
        count = df[df[4].str.contains(material, na=False)][3].values[0]
    except:
        count = 'NOT FOUND'

    return count



# Permutation to get all materials for all measures
def measure_materials(row):
    hhn = row.name[0]
    mn = row.name[1]
    
    try:
        df = pd.read_excel(f'Workbooks/{str(hhn)}', f'm{str(mn)}', header=None)
    except:
        df = pd.read_excel(f'Workbooks/{str(hhn)}', f'M{str(mn)}', header=None)
    
    materials_col = df[4, 2:20].values
    quantity_col = df[3, 2:20].values

    return materials_col, quantity_col


# Discern what excel files relate to the DF of work orders, then scrape the data from them and add that to a new list
relevant_hhns = []
result_count = []

material_term = 'CELLULOSE' # What material search for in the measure detail sheets

# Searching 
for hhn, m in measure_ids:
    relevant_wb = ''
    for wb in crew_workbooks:
        if wb.find(hhn) != -1:
            relevant_wb = wb
    relevant_hhns.append(str(hhn))
    result_count.append(measure_detail_scraper(relevant_wb, m, material_term))


# Add a new column to the measures DataFrame with the scraped material use data
base_with_materials_df[material_term+' Used'] = result_count






## Filter out unrelated workbooks
relevant_houses = []

# Get house data to help normalize measure hours spent per home's square footage
houses_df = pd.read_pickle('Output 4 - House Data/house_data.pkl')
relevant_houses = houses_df.loc[relevant_hhns]

houses_sqft = relevant_houses.loc[:, 'Sqft'].values

houses_sqft_mod = []

for sqft in houses_sqft:
    if sqft.isnumeric():
        houses_sqft_mod.append(int(sqft))
    else:
        houses_sqft_mod.append(None)

base_with_materials_df['SQFT'] = houses_sqft_mod


print(base_with_materials_df)






# Identify what DF to write to file
write_df = base_with_materials_df

# Write the constructed DataFrame to a .csv file for external use
write_df.to_csv(search_term + '.csv', index=True)
