import pandas as pd
import numpy as np
import os
import re
from PyPDF2 import PdfReader
import datetime as dt


input_folder = 'Old/NEAT Reports/'

# Load NEAT Reports
reports = os.listdir(input_folder)
print('NEAT Reports Found: ' + str(len(reports)))






# Find House IDs by regex patterns in workbook file names
hhn_list_by_pattern = []
hhn_errors = []

def hhn_checker(file_name):
    hhn_pattern = "\d{7}|\d{6}[abcd]|\d{6}|\d{5}"
    file_string = str(file_name).lower()
    check = re.search(hhn_pattern, file_string)
    if check != None:
        return re.findall(hhn_pattern, file_string)


# Search a NEAT report PDF for search terms, return the input value from the report
search_terms = ('FloorArea (sqft)', 'Year Built')

def term_searcher(search_term, file):
    reader1 = PdfReader(input_folder+str(file))
    page1 = reader1.pages[0]
    text1 = page1.extract_text()
    
    # Finding Year Built and Square Footage
    find = text1.find(search_term)
    mod = len(search_term)
    if search_term == 'Address':
        val = text1[find+mod:find+mod+50]
        eldint = str.find(val, 'Elderly')
        result = val[0:eldint]
    else:
        val = text1[find+mod:find+mod+5]
        result = re.sub("[^0-9]", "", val)

    return result


# Search a NEAT report for a house's ZIP
def zip_searcher(file):
    reader = PdfReader(input_folder+r)
    page = reader.pages[0]
    text1 = page.extract_text()
    find2 = text1.find('CityStateZip')
    val = text1[find2+13:find2+40]

    try:
        zip_check = re.search('[0-9]{5}', val)
        check = re.search('MN', val)
        span = check.span()
        return val[0:span[0]], zip_check.group()
    except AttributeError:
        return (None, None)


# Collect a file's creation date per the OS
wb_path = 'Workbooks/'
workbooks = os.listdir(wb_path)

def file_date(path, file):
    #secs = os.path.getctime(path+file)
    secs = os.path.getmtime(path+file)
    creation_date = dt.datetime.fromtimestamp(secs).strftime('%Y-%m-%d')
    return creation_date






houses_data = []

for r in reports:
    # Create a list for data on each home to be added to
    house_data = []
    
    # Find house's HHN in NEAT report using hhn_checker, append
    try:
        hhn = hhn_checker(r)[0].lower()
    except TypeError:
        hhn = None
        print(f'Check file: {r}')
    house_data.append(hhn)

    for term in search_terms:
        house_data.append((term_searcher(term, r)))

    geo = zip_searcher(r)

    #house_data.append(geo[0])
    house_data.append(geo[1])

    # Get and Append the date of wb creation to house data (more or less == date of work completion)
    for wb in workbooks:
        if wb.find(hhn) > 1:
            house_data.append(file_date(wb_path, wb))

    # For some reason the date adds twice to some workbooks, so if the list is too long pop trims it to proper size
    if len(house_data) > 5:
        house_data.pop()

    houses_data.append(house_data)

# Name columns, set index to House ID
houses_df = pd.DataFrame(houses_data, columns=('House ID', 'Sqft','Year Built', 'ZIP', 'WB Completion Date'))
houses_df.set_index('House ID', inplace=True)

year_test = houses_df['Year Built'].value_counts()
print(year_test)






# Import WA client & household data as a dataframe
wa_data = pd.read_csv('wa-client-house-info.csv')

# Subset of the whole WA dataframe with only relevant columns, rename columns, set index to House ID
wa_rel_data = wa_data[['ClientRecordName', 'ClientAddress', 'ClientCity', 'ClientCounty', 'ClientZip']]
wa_rel_data.columns = ['House ID', 'Address', 'City', 'County', 'ZIP']
wa_rel_data.set_index('House ID', inplace=True)

print(wa_rel_data.head())

# Merge WA data into the data collected from NEAT reports
# I really wanna make this so that it prioritizes WA data and backfills with NEAT data if WA data is missing
houses_df2 = houses_df.merge(wa_rel_data, left_index=True, right_index=True, how='left')

print(houses_df2)






# Add initial and final blower door results to house data
bd_df = pd.read_csv('Output 4 - House Data/hhns.csv', index_col=0)

houses_comb = pd.merge(houses_df, bd_df, left_index=True, right_index=True, how='outer')

print(houses_comb.head())






# Add crew assigned to each house
def get_crew_list(folder, file):
    df = pd.read_pickle(folder+file)
    crew_list = []
    pkl_hhn = str(file[:-4]).lower()

    for m in df['Crew']:
        crew_list.extend(m)
    try:
        crew_list.remove('')
    except ValueError:
        pass

    crew_sorted = list(set(crew_list))
    crew_sorted.sort()
    crew_counted = [crew for crew in crew_sorted if len(crew) > 0]

    return [pkl_hhn, crew_counted]


pickle_folder = 'Output 2 - Unify/Pickles/'
pickles = os.listdir(pickle_folder)

crew_per_house = []

for p in pickles:
    result = get_crew_list(pickle_folder, p)
    crew_per_house.append(result)

crew_per_house_df = pd.DataFrame(crew_per_house, columns=['House ID', 'Crew'])
crew_per_house_df.set_index('House ID', inplace=True)

houses_w_crew = houses_comb.merge(crew_per_house_df, left_index=True, right_index=True, how='outer')

houses_w_crew['ZIP'] = pd.to_numeric(houses_w_crew['ZIP'], errors='coerce')

print(houses_w_crew)






df_to_file = houses_w_crew

# Save the dataframe as a Pickle for internal convenience
df_to_file.to_pickle('Output 4 - House Data/house_data.pkl')

# Save the dataframe to a CSV for external convenience
df_to_file.to_csv('Output 4 - House Data/house_data.csv', index=True)
