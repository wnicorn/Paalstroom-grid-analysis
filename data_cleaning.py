"""
Author: Elise van Wijngaarden
Last updated: 18 june 2022

INSTRUCTIONS:
Before performing the Gaia simulations, please run the 'network_types.py'
to create the correct directories for saving the results.

Before running this script, the Gaia simulations must tbe run.

The variables under the 'Input' section right below needs to be updated in order to run this script:
- Edit 'results_dir' to the common folder containing all the result folders with the result excels.
- Edit 'network_files_dir' to the location where all the Gaia network files are.
- Edit 'language' to the language in which the Gaia results have been saved (English or Nederlands).

If the instructions above are met, this script saves a csv file with the cleaned results.
Afterwards, the Graphics script can be run.

(Time for teh first ~40 network files: 2116 sec > 35min run time)
"""

# ------------------- Input: -------------------------
# Common Gaia results folders location:
results_dir = 'C:/Users/Elise/Desktop/Gaia/'

# Input files location:
network_files_dir = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Technical Evaluation/Grid/Collected Data/Enexis/Round 2/additional' # TODO: remove afterwards

# language used in results excels:
language = 'English'  # 'English' or 'Nederlands'

# ---------------------------------------------------


import pandas as pd
import os
import time
import re

start = time.time()
print(time, 'Started data cleaning...')

# prep
engine = 'openpyxl'
if language == 'English':
    locations = ['Nodes', 'Branches', 'Elements', 'Switches and protections']  # , 'Hoofdkabels']
    loading = 'Load rate'  # 'Belastinggraad'
    voltages_phase = ['UL1', 'UL2', 'UL3']
    voltages_phase_n = ['UL1N', 'UL2N', 'UL3N']
    sort = 'Sort'
    trafo = 'transformer'
    cable = 'cable'
elif language == 'Nederlands':
    locations = ['Knooppunten', 'Takken', 'Elementen', 'Schakelaars en beveiligingen']  # , 'Hoofdkabels']
    loading = 'Belastinggraad'
    voltages_phase = ['UL1', 'UL2', 'UL3']
    voltages_phase_n = ['UL1N', 'UL2N', 'UL3N']
    sort = 'Soort'
    trafo = 'transformator'
    cable = 'kabel'
else:
    ValueError('Language not recognized')

include = voltages_phase + voltages_phase_n + [loading, sort]
aspects = ['Voltage', 'Loading', 'Fuse']

voltage_max = 253
voltage_min = 207

# load all network file names for identification
list_files = os.listdir(network_files_dir)
gnf_file_names = [filename[-14:-4] for filename in list_files if filename.endswith('.gnf')]

# start loop per network analyzed (folder of excel files from Gaia)
print(time, 'Starting counting violations...')
networks = {}
for network in gnf_file_names:
    network_folder = os.path.join(results_dir, network)
    list_files = os.listdir(network_folder)
    excel_files = [filename for filename in list_files if filename.endswith('.xlsx')]

    # define tests
    tests = {}  # later saved to networks[network]['tests']
    nr = 0
    for filename in excel_files:
        cable_group = re.match(r'(.+)_\d+_result_.+.xlsx$', filename).group(1)
        n_ov = re.match(r'.+_(\d+)_result_.+.xlsx$', filename).group(1)
        label = re.match(r'.+_result_(.+).xlsx$', filename).group(1)
        scenario = re.match(r'(.+)_\d+$', label).group(1)
        label_int = int(re.match(r'.+_(\d+)$', label).group(1))
        if label_int < 10:
            label_nr = '0' + str(label_int)
            new_label = label[:-1] + label_nr
        else:
            label_nr = label_int
            new_label = label

        tests[nr] = {'id': new_label, 'cable_group': cable_group, 'n_ov': n_ov,
                     'file': filename,
                     'light': re.match(r'^(.{3})_.+', new_label).group(1),
                     'charger': re.match(r'.+_(.+)_\d+$', new_label).group(1),
                     'nr': label_nr,
                     'Voltage_violations': {}, 'Loading_violations': {}, 'Fuse_violations': {},
                     'total_v_violations': 0, 'total_l_violations': 0, 'total_f_violations': 0}
        nr += 1

    # ------- Data Analysis -------

    # start loop per file (per test)
    networks[network] = {}
    violations = {}
    for test in tests.keys():
        file = tests[test]['file']
        scenario = re.match(r'.+_result_(.+)_\d+.xlsx$', file).group(1)

        # pre-allocation
        violations[aspects[0]] = {locations[0]: 0, locations[2]: 0}
        violations[aspects[1]] = {locations[1]: 0, locations[2]: 0, trafo: 0}
        violations[aspects[2]] = {locations[3]: 0}

        df = {}
        # Analysis per Excel sheet (location)
        for elem in locations:
            # import files as dataframes per sheet
            df[elem] = pd.read_excel(os.path.join(network_folder, file), engine=engine, sheet_name=elem, header=0,
                                     skiprows=[1])
            # only include relevant columns
            df[elem] = df[elem][[name for name in include if name in df[elem].columns]]
            # drop rows that contain NaN values because they are not useful
            df[elem] = df[elem].dropna()

            # check voltages at nodes
            if elem == locations[0]:
                for i in df[elem].index:
                    if all(x > 0 for x in df[elem][voltages_phase].loc[i].to_list()):
                        if any(x < 207.0 for x in df[elem][voltages_phase].loc[i].to_list()):
                            violations[aspects[0]][elem] += 1
                        # if single phase ?

            # check loading of cables and transformers
            if elem == locations[1]:
                for i in df[elem].index:
                    if df[elem].at[i, sort] == trafo:
                        if df[elem][loading].loc[i] > 100:
                            violations[aspects[1]][trafo] += 1
                    else:
                        if df[elem][loading].loc[i] > 100:
                            violations[aspects[1]][elem] += 1

            # check voltages and loadings at connections (elements)
            if elem == locations[2]:
                for i in df[elem].index:
                    if all(x > 0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                        if any(x < 207.0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                            violations[aspects[0]][elem] += 1
                        # if single phase ?
                    if df[elem][loading].loc[i] > 100:
                        violations[aspects[1]][elem] += 1

            # check fuses
            if elem == locations[3]:
                for i in df[elem].index:
                    if df[elem][loading].loc[i] > 100:
                        violations[aspects[2]][elem] += 1

        # store split and total values
        for x in aspects:
            tests[test]['{}_violations'.format(x)] = violations[x]
            tests[test]['total_{}_violations'.format(x[0])] = sum(violations[x].values())

    networks[network]['tests'] = tests
    print(time, ' Done:', network)

# ------- data aggregation -------

active_networks = [x for x in networks.keys() if networks[x]['tests']]
tests = networks[active_networks[0]]['tests']
columns = []
for x in aspects:
    for a in tests[0]['{}_violations'.format(x)]:
        columns.append('{}'.format(a))
    columns.append('Total')

print(time, 'Starting combining data...')
dfs = {}
for network in active_networks:
    tests = networks[network]['tests']
    index_nr = [tests[a]['id'] for a in tests.keys()]
    df = pd.DataFrame(index=index_nr, columns=columns)
    # create a multi-index for nicer graph indexes:
    header = [aspects[0], aspects[0], aspects[0], aspects[1], aspects[1], aspects[1], aspects[1], aspects[2],aspects[2]]
    df.columns = pd.MultiIndex.from_tuples(list(zip(header, df.columns)), names=['Type', 'Scenario'])

    chars = {}
    for char in list(tests[0].keys())[1:3]:
        chars[char] = [tests[a][char] for a in tests.keys()]

    for i in tests.keys():
        data = []
        for x in aspects:
            for elem in tests[i]['{}_violations'.format(x)].keys():
                df.loc[tests[i]['id'], ('{}'.format(x), '{}'.format(elem))] = tests[i]['{}_violations'.format(x)][
                    '{}'.format(elem)]
            df.loc[tests[i]['id'], ('{}'.format(x), 'Total')] = tests[i]['total_{}_violations'.format(x[0])]

    df.index = [list(range(len(df.index))), [network for n in df.index], chars['cable_group'], chars['n_ov'], df.index.str[:3], df.index.str[4:-3], df.index.str[-2:]]
    df.rename_axis(['index', 'Network', 'Cable group', 'Connections', 'Light type', 'Charger power', 'Location'], inplace=True)
    df.sort_index()

    dfs[network] = df
    print(time, ' Done:', network)
df_combined = pd.DataFrame()
for df in dfs.values():
    df_combined = pd.concat([df_combined, df])


# Save numeric values in dataframe to csv:
data = df_combined.apply(pd.to_numeric)
# data.rename(columns={'Switches and protections': 'Feeder', trafo: 'Transformer'})
filename = 'cleaned_results_2.csv'
data.to_csv(os.path.join(results_dir, filename))
print('Saved results to:', filename)

end = time.time()
time = end - start
print('All done! \ntime: ', round(time/60, 2), 'min')
