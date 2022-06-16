'''
Author: Elise van Wijngaarden
Last updated: 16 june 2022

INSTRUCTIONS:
Before performing the Gaia simulations, please run the 'network_types.py'
to create the correct directories for saving the results.

Before running this script, the Gaia simulations must tbe run.

If the instructions above are met, only the 'Input' here at the top needs to be updated in order to run this script.

'''

# ------------------- Input: -------------------------
# Gaia results folders location:
common_directory = 'C:/Users/Elise/Desktop/Gaia/'

# characteristics file:
characteristics_dir = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Technical Evaluation/Grid/Collected Data/Enexis/Round 2/'
characteristics_excel_name = '20220307_Results_SL_R2102-02_gG-Light.xlsx'

# language used in results excels:
language = 'English'  # 'English' or 'Nederlands'

# ---------------------------------------------------


import pandas as pd
import os
import time
import re
import numpy as np
import matplotlib.pyplot as plt
from itertools import groupby
import seaborn as sns

start = time.time()

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

# load characteristics data
list_files = os.listdir(characteristics_dir)
gnf_file_names = [filename[-14:-4] for filename in list_files if filename.endswith('.gnf')]

char_file = os.path.join(characteristics_dir, characteristics_excel_name)
char_df = pd.read_excel(char_file, sheet_name='Steekproef', engine='openpyxl', header=0)
char_df = char_df[['KABELGROEP', 'NETSTATION', 'NETSTATION_STEDELIJKHEID', 'NETWERKTYPE', 'KABELGROEP_AANLEGJAAR', 'KABELGROEP_DECENIUM',
         'KABELGROEP_LENGTE', 'NODE_MAX_FOUTSPANNING', 'N_OV_AANSLUITING']]

# start loop per network analyzed (folder of excel files from Gaia)
networks = {}
for network in gnf_file_names:
    network_folder = os.path.join(common_directory, network)
    list_files = os.listdir(network_folder)
    excel_files = [filename for filename in list_files if filename.endswith('.xlsx')]

    # define tests
    tests = {}  # later saved to networks[network]['tests']
    nr = 0
    for filename in excel_files:
        cable_group = re.match(r'(.+)_result_.+.xlsx$', filename).group(1)
        label = re.match(r'.+_result_(.+).xlsx$', filename).group(1)
        scenario = re.match(r'(.+)_\d+$', label).group(1)
        label_int = int(re.match(r'.+_(\d+)$', label).group(1))
        if label_int < 10:
            label_nr = '0' + str(label_int)
            new_label = label[:-1] + label_nr
        else:
            label_nr = label_int
            new_label = label

        tests[nr] = {'id': new_label, 'cable_group': cable_group,
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


# ------- data aggregation -------

active_networks = [x for x in networks.keys() if networks[x]['tests']]
tests = networks[active_networks[0]]['tests']
columns = []
for x in aspects:
    for a in tests[0]['{}_violations'.format(x)]:
        if a == 'Switches and protections':
            columns.append('Feeders')
        elif a == 'transformer':
            columns.append('Transformer')
        else:
            columns.append('{}'.format(a))
    columns.append('Total')

dfs = {}
for network in active_networks:
    tests = networks[network]['tests']
    index_labels = [tests[a]['id'] for a in tests.keys()]
    df = pd.DataFrame(index=index_labels, columns=columns)
    # create a multi-index for nicer graph indexes:
    header = [aspects[0], aspects[0], aspects[0], aspects[1], aspects[1], aspects[1], aspects[1], aspects[2], aspects[2]]
    df.columns = pd.MultiIndex.from_tuples(list(zip(header, df.columns)), names=['Type', 'Location'])

    for i in tests.keys():
        data = []
        for x in aspects:
            for elem in tests[i]['{}_violations'.format(x)].keys():
                df.loc[tests[i]['id'], ('{}'.format(x), '{}'.format(elem))] = tests[i]['{}_violations'.format(x)][
                    '{}'.format(elem)]
            df.loc[tests[i]['id'], ('{}'.format(x), 'Total')] = tests[i]['total_{}_violations'.format(x[0])]

    df.index = [[network for l in df.index], df.index.str[:3], df.index.str[4:-3], df.index.str[-2:]]
    df.rename_axis(['Network', 'Light type', 'Charger power', 'Location'], inplace=True)
    df.sort_index()

    dfs[network] = df
df_combined = pd.DataFrame()
for df in dfs.values():
    df_combined = pd.concat([df_combined, df])


# aggregate to violation types for a test type
df = df_combined.apply(pd.to_numeric)
totals = df.loc[:, df.columns.get_level_values(1) == 'Total']
specifics = df.loc[:, df.columns.get_level_values(1) != 'Total']

tot_light_charger_mean = totals.groupby(level=['Light type', 'Charger power']).mean()
spe_light_charger_mean = specifics.groupby(level=['Light type', 'Charger power']).mean()

totals_led = df.loc[df.index.get_level_values('Light type') == 'LED', df.columns.get_level_values(1) == 'Total']
specifics_led = df.loc[df.index.get_level_values('Light type') == 'LED', df.columns.get_level_values(1) != 'Total']

# tot_led_cat1_charger_mean = totals_led.groupby(level=[cat_1, 'Charger power']).mean()
# spe_led_cat1_charger_mean = specifics_led.groupby(level=[cat_1, 'Charger power']).mean()

# Print numerical results
df_nr_results = pd.DataFrame(index=['tests', 'violations', 'ratio [%]'])

df_nr_results['Total'] = [int(len(df)), int(len(df) - len(df[df[df.columns] == 0].dropna(inplace=False))), round((len(df) - len(df[df[df.columns] == 0].dropna(inplace=False))) / len(df) * 100, 2)]

data = df.loc[df.index.get_level_values('Light type') == 'LED']
if len(data) > 0:
    df_nr_results['LED'] = [int(len(data)), int(len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))),
                            round((len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))) / len(data) *100, 2)]

data = df.loc[df.index.get_level_values('Light type') == 'old']
if len(data) > 0:
    df_nr_results['old'] = [int(len(data)), int(len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))),
                            round((len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))) / len(data) *100, 2)]

# data = df.loc[df.index.get_level_values(cat_2) == 'Small']
# if len(data) > 0:
#     df_nr_results['Small'] = [int(len(data)), int(len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))),
#                             round((len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))) / len(data) *100, 2)]
#
# data = df.loc[df.index.get_level_values(cat_2) == 'Large']
# if len(data) > 0:
#     df_nr_results['Large'] = [int(len(data)), int(len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))),
#                             round((len(data) - len(data[data[data.columns] == 0].dropna(inplace=False))) / len(data) *100, 2)]


# -------------- Graphs --------------
def get_pie_labels(labels, sizes):
    labels = [f'{l}, {s:0.1f}%' for l, s in zip(labels, sizes)]
    plt.legend(bbox_to_anchor=(0.85, 1), loc='upper left', labels=labels)


# Plot pie chart % of general types of violations
fig1, ax1 = plt.subplots(figsize=(6, 4))
ax1.pie(totals.sum(axis=0), autopct='%1.1f%%', pctdistance=0.85)
plt.title('General violation type occurances')
plt.legend(list(totals.columns.get_level_values(0)))
plt.show()
# Plot pie chart % of specific types of violations
fig4, ax4 = plt.subplots(figsize=(6, 4))
plt.pie(specifics.sum(axis=0), autopct='%1.1f%%', pctdistance=0.85)
plt.title('Specific violation type occurances')
plt.legend(list(specifics.columns.get_level_values(0)))
plt.show()


# define operations to add category lines to bar plots:
def add_line(ax, xpos, ypos):
    line = plt.Line2D([xpos, xpos], [ypos + .095, ypos],
                      transform=ax.transAxes, color='black')
    line.set_clip_on(False)  # [... + .1, ...]
    ax.add_line(line)


def label_len(my_index, level):
    labels = my_index.get_level_values(level)
    return [(k, sum(1 for i in g)) for k, g in groupby(labels)]


def label_group_bar_table(ax, fig, df):
    labels = ['' for item in ax.get_xticklabels()]
    ax.set_xticklabels(labels)
    ypos = -.1
    scale = 1. / df.index.size
    for level in range(df.index.nlevels)[::-1]:
        pos = 0
        for label, rpos in label_len(df.index, level):
            lxpos = (pos + .5 * rpos) * scale
            ax.text(lxpos, ypos, label, ha='center', transform=ax.transAxes)
            add_line(ax, pos * scale, ypos)
            pos += rpos
        add_line(ax, pos * scale, ypos)
        ypos -= .1
    ax.xaxis.set_label_coords(0.5, -0.25)
    fig.subplots_adjust(bottom=.1 * df.index.nlevels)


# plot bar plot mean # general violations types per scenario type
data = tot_light_charger_mean
fig2, ax2 = plt.subplots(figsize=(6, 4))
data.set_axis(data.columns.map(', '.join), axis=1, inplace=False).plot(kind='bar', stacked=True, ax=ax2)
plt.title('Average general violation types per scenario')
label_group_bar_table(ax2, fig2, tot_light_charger_mean)
plt.show()

# plot bar plot mean # specific violations types per scenario type
data = spe_light_charger_mean
fig3, ax3 = plt.subplots(figsize=(6, 4))
data.set_axis(data.columns.map(', '.join), axis=1, inplace=False).plot(kind='bar',
                                                                                                           stacked=True,
                                                                                                           ax=ax3)
plt.title('Average specific violation types per scenario')
label_group_bar_table(ax3, fig3, spe_light_charger_mean)
plt.show()

# # plot bar plot mean # general violations types per network type for LED scenario
# fig5, ax5 = plt.subplots(figsize=(6, 4))
# tot_led_cat1_charger_mean.set_axis(tot_led_cat1_charger_mean.columns.map(', '.join), axis=1, inplace=False).plot(
#     kind='bar', stacked=True, ax=ax5)
# plt.title('Average general violation types per scenario')
# label_group_bar_table(ax5, fig5, tot_led_cat1_charger_mean)
# plt.show()
# # plot bar plot mean # specific violations types per network type for LED scenario
# fig6, ax6 = plt.subplots(figsize=(6, 4))
# spe_led_cat1_charger_mean.set_axis(spe_led_cat1_charger_mean.columns.map(', '.join), axis=1, inplace=False).plot(
#     kind='bar', stacked=True, ax=ax6)
# plt.title('Average specific violation types per scenario')
# label_group_bar_table(ax6, fig6, spe_led_cat1_charger_mean)
# plt.show()

end = time.time()
time = end - start
print('time: ', round(time, 2), 'sec')
