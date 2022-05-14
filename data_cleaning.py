import pandas as pd
import os
import time

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

voltage_max = 253
voltage_min = 207

cwd = os.getcwd()
all_files = os.listdir(cwd)
excel_files = [filename for filename in all_files if filename.endswith('fuses.xlsx')]

# define scenarios
scenarios = {}
nr = 1
for filename in excel_files:
    scenarios[nr] = {'id': nr, 'file': filename,
                     'voltage_violations': {}, 'loading_violations': {}, 'fuse_violations': {},
                     'total_v_violations': 0, 'total_l_violations': 0, 'total_f_violations': 0,
                     }
    nr += 1

# start loop per file (scenario)
for scenario in scenarios.keys():
    entry = scenarios[scenario]
    file = entry['file']

    # pre-allocation
    voltage_violations = {}
    loading_violations = {trafo: 0}
    fuse_violations = {}

    df = {}
    # Analysis per Excel sheet
    for elem in sheets:
        # import files as dataframes per sheet
        df[elem] = pd.read_excel(file, engine=engine, sheet_name=elem, header=0, skiprows=[1])
        # only include relevant columns
        df[elem] = df[elem][[name for name in include if name in df[elem].columns]]
        # drop rows that contain NaN values because they are not useful
        df[elem] = df[elem].dropna()
        # initialize counting
        voltage_violations[elem] = 0
        loading_violations[elem] = 0
        fuse_violations[elem] = 0

        # check voltages at nodes
        if elem == sheets[0]:
            for i in df[elem].index:
                if all(x > 0 for x in df[elem][voltages_phase].loc[i].to_list()):
                    if any(x < 207.0 for x in df[elem][voltages_phase].loc[i].to_list()):
                        voltage_violations[elem] += 1
                    # if single phase ?

        # check loading of cables and transformers
        if elem == sheets[1]:
            for i in df[elem].index:
                if df[elem].at[i, sort] == trafo:
                    if df[elem][loading].loc[i] > 100:
                        loading_violations[trafo] += 1
                else:
                    if df[elem][loading].loc[i] > 100:
                        loading_violations[elem] += 1

        # check voltages and loadings at connections (elements)
        if elem == sheets[2]:
            for i in df[elem].index:
                if all(x > 0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                    if any(x < 207.0 for x in df[elem][voltages_phase_n].loc[i].to_list()):
                        voltage_violations[elem] += 1
                    # if single phase ?
                if df[elem][loading].loc[i] > 100:
                    loading_violations[elem] += 1

        # check fuses
        if elem == sheets[3]:
            for i in df[elem].index:
                if df[elem][loading].loc[i] > 100:
                    fuse_violations[elem] += 1

    entry['voltage_violations'] = voltage_violations
    entry['loading_violations'] = loading_violations
    entry['fuse_violations'] = fuse_violations
    entry['total_v_violations'] = sum(voltage_violations.values())
    entry['total_l_violations'] = sum(loading_violations.values())
    entry['total_f_violations'] = sum(fuse_violations.values())

    # print results
    print('Scenario: ', entry['id'])
    print('voltage violations: ', voltage_violations)
    print('loading violations: ', loading_violations)
    print('fuse violations: ', fuse_violations)
    print('total voltage violations: ', entry['total_v_violations'])
    print('total loading violations: ', entry['total_l_violations'])
    print('total fuse violations: ', entry['total_f_violations'])

end = time.time()
time = end-start
print('time: ', round(time, 4))

# --- Graphics ---


