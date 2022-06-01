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
language = 'English' # 'Nederlands'
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
# nr = 1
for filename in excel_files:
    label = re.match(r'result_(.+).xlsx$', filename).group(1)
    scenario = re.match(r'(.+)_\d+$', label).group(1)
    nr = re.match(r'_(\d+)$', label).group(1)
    # TODO: FIX!
    # tests[nr] = {'id': label, 'file': filename,
    #              'light': re.match(r'^(\s{3})', label).group(1),
    #              'charger': re.match(r'_(.+)_\d+$', label).group(1),
    #              'voltage_violations': {}, 'loading_violations': {}, 'fuse_violations': {},
    #              'total_v_violations': 0, 'total_l_violations': 0, 'total_f_violations': 0}
    # nr += 1

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
        df[elem] = pd.read_excel(file, engine=engine, sheet_name=elem, header=0, skiprows=[1])
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
for i in range(len(tests)):  # test keys are integers from [1, ...]
    data.append([])
    for x in aspects:
        bs = tests[i+1]['{}_violations'.format(x)].values()
        tots = tests[i+1]['total_{}_violations'.format(x[0])]
        data[i].extend(bs)
        data[i].append(tots)
df = pd.DataFrame(index=tests.keys(), columns=columns, data=data)

# create a multi-index for nicer graph indexes:
header = [aspects[0], aspects[0], aspects[0], aspects[1], aspects[1], aspects[1], aspects[2], aspects[2]]
df.columns = pd.MultiIndex.from_tuples(list(zip(header, df.columns)), names=['Type', 'Location'])
df.index.names = ['test']
#   df.columns.set_levels(['b1','c1','f1'],level=1,inplace=True)
# print(df.columns)
print(df)

# violation types
triggers = {}
for x in aspects:
    triggers[x] = 0
    for n in tests.values():
        if sum(n['{}_violations'.format(x)].values()) > 0:
            triggers[x] += 1
    # print('# tests with violated {}s:'.format(x), triggers[x])



# ------- Graphics -------

# Plot pie chart % of types of violations
fig1, ax1 = plt.subplots()
ax1.pie(triggers.values(), labels=triggers.keys(), autopct='%1.1f%%') #, startangle=90)
plt.title('Violation type occurances')
# plt.show()


# plot total stacked bar graph of totals
fig2, ax2 = plt.subplots()
df.loc[:, df.columns.get_level_values(1) == 'Total'].plot(kind='bar', stacked=True, ax=ax2)
plt.show()


fig3, ax3 = plt.subplots()
df.loc[:, df.columns.get_level_values(1) != 'Total'].plot(kind='bar', stacked=True, ax=ax3)
plt.show()

# aggregate to violation types for a test type

test = 'P = x'
start = 0
end = 2
selection = df.loc[start:end, df.columns.get_level_values(1) != 'Total']
fig4, ax4 = plt.subplots()
selection.mean().unstack().plot(kind='bar', stacked=True, ax=ax4)
plt.title('Average violation types for tests {}'.format(test))
plt.show()







end = time.time()
time = end-start
# print('time: ', round(time, 4))

