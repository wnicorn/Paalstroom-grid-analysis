import pandas as pd
import os
import time
import re
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

start = time.time()

# prep
engine = 'openpyxl'
language = 'Nederlands' # 'Nederlands'
if language == 'English':
    locations = ['Nodes', 'Branches', 'Elements', 'Switches and protections'] #, 'Hoofdkabels']
    loading = 'Load rate' #'Belastinggraad'
    voltages_phase = ['UL1', 'UL2', 'UL3']
    voltages_phase_n = ['UL1N', 'UL2N', 'UL3N']
    sort = 'Sort'
    trafo = 'transformer'
    cable = 'cable'
elif language == 'Nederlands':
    locations = ['Knooppunten', 'Takken', 'Elementen', 'Schakelaars en beveiligingen'] #, 'Hoofdkabels']
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

# cwd = os.getcwd()
loc = 'C:/Users/Elise/Desktop/Gaia/'
all_files = os.listdir(loc)
excel_files = [filename for filename in all_files if filename.endswith('.xlsx')]

# define scenarios
scenarios_listed = []
nrs_listed = []
for filename in excel_files:
    scenarios_listed.append(re.match(r'result_(.+)_\d+\.xlsx$', filename).group(1))
    nrs_listed.append(re.match(r'result_.+_(\d+)\.xlsx$', filename).group(1))
scenario_types = np.unique(np.array(scenarios_listed))
nrs = np.sort(np.unique(np.array(nrs_listed)).astype(int))
scenarios = {}

# define tests
tests = {}
violations = {}
nr = 0
for filename in excel_files:
    label = re.match(r'result_(.+).xlsx$', filename).group(1)
    scenario = re.match(r'(.+)_\d+$', label).group(1)
    label_int = int(re.match(r'.+_(\d+)$', label).group(1))
    if label_int < 10:
        label_nr = '0'+str(label_int)
        new_label = label[:-1]+label_nr
    else:
        label_nr = label_int
        new_label = label

    tests[nr] = {'id': new_label, 'file': filename,
                 'light': re.match(r'^(.{3})_.+', new_label).group(1),
                 'charger': re.match(r'.+_(.+)_\d+$', new_label).group(1),
                 'nr': label_nr,
                 'voltage_violations': {}, 'loading_violations': {}, 'fuse_violations': {},
                 'total_v_violations': 0, 'total_l_violations': 0, 'total_f_violations': 0}
    nr += 1

# ------- Data Analysis -------

scenarios = {}

# start loop per file (test)
for test in tests.keys():
    entry = tests[test]
    file = entry['file']
    scenario = re.match(r'result_(.+)_\d+.xlsx$', file).group(1)

    # pre-allocation
    violations[aspects[0]] = {locations[0]: 0, locations[2]: 0}
    violations[aspects[1]] = {locations[1]: 0, trafo: 0}
    violations[aspects[2]] = {locations[3]: 0}

    df = {}
    # Analysis per Excel sheet (location)
    for elem in locations:
        # import files as dataframes per sheet
        df[elem] = pd.read_excel(os.path.join(loc, file), engine=engine, sheet_name=elem, header=0, skiprows=[1])
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
        entry['{}_violations'.format(x)] = violations[x]
        entry['total_{}_violations'.format(x[0])] = sum(violations[x].values())

    # print results
    # print('test: ', entry['id'])
    # for x in aspects:
    #     print('{} violations: '.format(x))#, violations[x])
    #     for a in violations[x].keys():
    #         print(' ', a,':', violations[x][a])
    #     print('total {} violations: '.format(x), entry['total_{}_violations'.format(x[0])])



# ------- data aggregation -------
#   print(tests.values())

columns = []  # '{}_violations'.format(x) for x in aspects
for x in aspects:
    for a in tests[test]['{}_violations'.format(x)]:
        if a == 'Switches and protections':
            columns.append('Feeders')
        elif a == 'transformer':
            columns.append('Transformer')
        else:
            columns.append('{}'.format(a))
    columns.append('Total')
#data = [[tests[i]['total_{}_violations'.format(x[0])] for x in aspects] for i in tests.keys()]

data = []
for i in range(len(tests)):  # test keys are integers from [0, ...]
    data.append([])
    for x in aspects:
        bs = tests[i]['{}_violations'.format(x)].values()
        tots = tests[i]['total_{}_violations'.format(x[0])]
        data[i].extend(bs)
        data[i].append(tots)
index_labels = [tests[a]['id'] for a in range(len(tests))]
df = pd.DataFrame(index=index_labels, columns=columns, data=data)

# create a multi-index for nicer graph indexes:
header = [aspects[0], aspects[0], aspects[0], aspects[1], aspects[1], aspects[1], aspects[2], aspects[2]]
df.columns = pd.MultiIndex.from_tuples(list(zip(header, df.columns)), names=['Type', 'Location'])
df.index = [df.index.str[:3], df.index.str[4:-3], df.index.str[-2:]]
df.rename_axis(['Light type', 'Charger power', 'Location'], inplace=True)
df.sort_index()
print(df)

# aggregate to violation types for a test type
totals = df.loc[:, df.columns.get_level_values(1) == 'Total']
specifics = df.loc[:, df.columns.get_level_values(1) != 'Total']
mean_totals = totals.groupby(level=['Light type', 'Charger power']).mean()
mean_specifics = specifics.groupby(level=['Light type', 'Charger power']).mean()
sum_totals = totals.groupby(level=['Light type', 'Charger power']).sum()
sum_specifics = specifics.groupby(level=['Light type', 'Charger power']).sum()


# ------- Graphics -------

# Plot pie chart % of types of violations
fig1, ax1 = plt.subplots()
ax1.pie(sum_totals.sum(axis=0), labels=list(sum_totals.columns.get_level_values(0)), autopct='%1.1f%%') #, startangle=90)
plt.title('Violation type occurances')
plt.show()

# plot bar plot mean # general violations types per scenario type
fig2, ax2 = plt.subplots()
mean_totals.plot(kind='bar', stacked=True, ax=ax2)
plt.title('Average general violation types per scenario')
plt.show()

# plot bar plot mean # specific violations types per scenario type
fig3, ax3 = plt.subplots()
mean_specifics.plot(kind='bar', stacked=True, ax=ax3)
plt.title('Average specific violation types per scenario')
plt.show()






end = time.time()
time = end-start
print('time: ', round(time, 4))

