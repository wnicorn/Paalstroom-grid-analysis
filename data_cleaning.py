import pandas as pd
import os
import time
import matplotlib.pyplot as plt

start = time.time()

# prep
engine = 'openpyxl'
language = 'English' # 'Nederlands'
if language == 'English':
    sheets = ['Nodes', 'Branches', 'Elements', 'Switches and protections'] #, 'Hoofdkabels']
    loading = 'Load rate' #'Belastinggraad'
    voltages_phase = ['UL1', 'UL2', 'UL3']
    voltages_phase_n = ['UL1N', 'UL2N', 'UL3N']
    sort = 'Sort'
    trafo = 'transformer'
    cable = 'cable'
elif language == 'Nederlands':
    sheets = ['Knooppunten', 'Takken', 'Elementen', 'Schakelaars en beveiligingen'] #, 'Hoofdkabels']
    loading = 'Belastinggraad'
    voltages_phase = ['UL1', 'UL2', 'UL3']
    voltages_phase_n = ['UL1N', 'UL2N', 'UL3N']
    sort = 'Soort'
    trafo = 'transformator'
    cable = 'kabel'
else:
    ValueError('Language not recognized')

include = voltages_phase + voltages_phase_n + [loading, sort]
aspects = ['voltage', 'loading', 'fuse']

voltage_max = 253
voltage_min = 207

cwd = os.getcwd()
all_files = os.listdir(cwd)
excel_files = [filename for filename in all_files if filename.endswith('fuses.xlsx')]

# define scenarios
scenarios = {}
violations = {}
nr = 1
for filename in excel_files:
    scenarios[nr] = {'id': nr, 'file': filename,
                     'voltage_violations': {}, 'loading_violations': {}, 'fuse_violations': {},
                     'total_v_violations': 0, 'total_l_violations': 0, 'total_f_violations': 0,
                     }
    nr += 1

# ------- Data Analysis -------
# start loop per file (scenario)
for scenario in scenarios.keys():
    entry = scenarios[scenario]
    file = entry['file']

    # pre-allocation
    violations['voltage'] = {sheets[0]: 0, sheets[2]: 0}
    violations['loading'] = {sheets[1]: 0, trafo: 0}
    violations['fuse'] = {sheets[3]: 0}

    df = {}
    # Analysis per Excel sheet
    for elem in sheets:
        # import files as dataframes per sheet
        df[elem] = pd.read_excel(file, engine=engine, sheet_name=elem, header=0, skiprows=[1])
        # only include relevant columns
        df[elem] = df[elem][[name for name in include if name in df[elem].columns]]
        # drop rows that contain NaN values because they are not useful
        df[elem] = df[elem].dropna()

        # check voltages at nodes
        if elem == sheets[0]:
            for i in df[elem].index:
                if all(x > 0 for x in df[elem][voltages_phase].loc[i].to_list()):
                    if any(x < 207.0 for x in df[elem][voltages_phase].loc[i].to_list()):
                        violations['voltage'][elem] += 1
                    # if single phase ?

        # check loading of cables and transformers
        if elem == sheets[1]:
            for i in df[elem].index:
                if df[elem].at[i, sort] == trafo:
                    if df[elem][loading].loc[i] > 100:
                        violations['loading'][trafo] += 1
                else:
                    if df[elem][loading].loc[i] > 100:
                        violations['loading'][elem] += 1

        # check voltages and loadings at connections (elements)
        if elem == sheets[2]:
            for i in df[elem].index:
                if all(x > 0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                    if any(x < 207.0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                        violations['voltage'][elem] += 1
                    # if single phase ?
                if df[elem][loading].loc[i] > 100:
                    violations['loading'][elem] += 1

        # check fuses
        if elem == sheets[3]:
            for i in df[elem].index:
                if df[elem][loading].loc[i] > 100:
                    violations['fuse'][elem] += 1

    # store split and total values
    for x in aspects:
        entry['{}_violations'.format(x)] = violations[x]
        entry['total_{}_violations'.format(x[0])] = sum(violations[x].values())

    # print results
    # print('Scenario: ', entry['id'])
    # for x in aspects:
        # print('{} violations: '.format(x), violations[x])
        # print('total {} violations: '.format(x), entry['total_{}_violations'.format(x[0])])




# ------- data aggregation -------
# print(scenarios.values())

columns = ['{}_violations'.format(x) for x in aspects]
for x in aspects:
    for a in scenarios[scenario]['{}_violations'.format(x)]:
        columns.append('{}_violations_at_{}'.format(x, a))
data = [[scenarios[i]['total_{}_violations'.format(x[0])] for x in aspects] for i in scenarios.keys()]
for i in range(len(data)):
    for x in aspects:
        for a in scenarios[i + 1]['{}_violations'.format(x)].values():
            data[i].append(a)
df = pd.DataFrame(index=scenarios.keys(), columns=columns, data=data)
print(df)

# violation types
triggers = {}
for x in aspects:
    triggers[x] = 0
    for n in scenarios.values():
        if sum(n['{}_violations'.format(x)].values()) > 0:
            triggers[x] += 1
    # print('# scenarios with violated {}s:'.format(x), triggers[x])



# ------- Graphics -------

# Plot pie chart % of types of violations
fig1, ax1 = plt.subplots()
ax1.pie(triggers.values(), labels=triggers.keys(), autopct='%1.1f%%') #, startangle=90)
# plt.show()

end = time.time()
time = end-start
# print('time: ', round(time, 4))

