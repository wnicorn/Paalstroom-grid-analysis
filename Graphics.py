import pandas as pd
import os
import math
import numpy as np
import time
import matplotlib.pyplot as plt
from itertools import groupby

"""
Author: Elise van Wijngaarden
Last updated: 21 june 2022

INSTRUCTIONS:
Before performing the Gaia simulations, please run the 'network_types.py'
to create the correct directories for saving the results.

Before running this script, the Gaia simulations and the data_cleaning.py script must be run.

The variables under the 'Input' section right below needs to be updated in order to run this script:
- Edit 'results_dir' to the common folder containing all the result folders with the result excels.
- Edit 'char_dir' to the location where the Gaia characteristics file is.
- Edit 'char_excel_name' only if the filename has changed.

If the instructions above are met, this script outputs three diagrams in the IDE.

"""

# ------------------- Input: -------------------------
# Common Gaia results folders location:
results_dir = 'C:/Users/Elise/Desktop/Gaia/'
results_dir_2 = 'C:/Users/Elise/Desktop/Gaia/additional'  # TODO: remove afterwards

# characteristics file:
char_dir = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Technical Evaluation/Grid/Collected Data/Enexis/Round 2/'
char_excel_name = '20220307_Results_SL_R2102-02_gG-Light.xlsx'

# ---------------------------------------------------

# start = time.time()

# import results file
filename = 'cleaned_results.csv'
filename_2 = 'cleaned_results_2.csv'  # TODO: remove afterwards
df_results = pd.read_csv(os.path.join(results_dir, filename), index_col=[0, 1, 2, 3, 4, 5, 6], header=[0, 1])
df_results_2 = pd.read_csv(os.path.join(results_dir, filename_2), index_col=[0, 1, 2, 3, 4, 5, 6], header=[0, 1])  # TODO: remove afterwards
df_results = pd.concat([df_results, df_results_2])  # TODO: remove afterwards

# fix some things...:
df_results = df_results.droplevel(level=0, axis=0).reset_index(col_level=1, col_fill='Scenario').sort_index(axis=0)
df_results = df_results.apply(pd.to_numeric, errors='ignore')
df_results.index.get_level_values(0).name = 'index'
df_results.columns.get_level_values(1).name = 'Specific'
df_results = df_results.drop(columns=[('Loading', 'Transformer'), ('Fuse', 'Feeders')])
df_results = df_results.rename(columns={'Switches and protections': 'Feeder', 'transformer': 'Transformer'})
df_results[('Scenario', 'Charger power')] = df_results[('Scenario', 'Charger power')].astype(str).replace(['_', '\.0'], '', regex=True).astype(float)
# import Enexis characteristics
df_char = pd.read_excel(os.path.join(char_dir, char_excel_name), engine='openpyxl', sheet_name='SL_gG_Cablegroup',
                        header=[0])
df_char = df_char[['KABELGROEP', 'NETSTATION_VESTIGING', 'NETSTATION_STEDELIJKHEID', 'NETWERKTYPE',
                   'KABELGROEP_DECENIUM', 'KABELGROEP_LENGTE', 'NODE_MAX_FOUTSPANNING', 'N_OV_AANSLUITING']
].dropna(how='all')  # 'NETSTATION', 'KABELGROEP_AANLEGJAAR',
df_char = df_char.loc[df_char['KABELGROEP'].isin(df_results[('Scenario', 'Cable group')])]
df_char['KABELGROEP_LENGTE_TYPE'] = df_char['KABELGROEP_LENGTE']
df_char['FOUTSPANNING_TYPE'] = df_char['NODE_MAX_FOUTSPANNING']
df_char['N_OV_AFGEROND'] = df_char['N_OV_AANSLUITING']
for i in df_char.index:
    df_char['KABELGROEP_LENGTE_TYPE'].loc[i] = math.ceil(df_char['KABELGROEP_LENGTE'].loc[i].astype(float) / 300) * 300
    df_char['N_OV_AFGEROND'].loc[i] = math.ceil(df_char['N_OV_AANSLUITING'].loc[i].astype(float) / 5) * 5
    df_char['FOUTSPANNING_TYPE'].loc[i] = math.floor(df_char['NODE_MAX_FOUTSPANNING'].loc[i].astype(float) / 50) * 50
    if df_char['NODE_MAX_FOUTSPANNING'].loc[i] >= 50:
        df_char['FOUTSPANNING_TYPE'].loc[i] = 'SWAP'
    else:
        df_char['FOUTSPANNING_TYPE'].loc[i] = 'single'
df_char.drop(['KABELGROEP_LENGTE', 'NODE_MAX_FOUTSPANNING', 'N_OV_AANSLUITING'], axis=1, inplace=True)
df_char.reset_index(inplace=True)
translate_urban = {
    'Niet_stedelijk': 'Not urban',
    'Weinig_stedelijk': 'Weakly urban',
    'Matig_stedelijk': 'Medium urban',
    'Sterk_stedelijk': 'Strongly urban',
    'Zeer_sterk_stedelijk': 'Very strongly urban'
}
translate_columns = {
    'NETSTATION_STEDELIJKHEID': 'Urbanism',
    'NETWERKTYPE': 'Network type',
    'KABELGROEP_DECENIUM': 'Decade',
    'KABELGROEP_LENGTE_TYPE': 'Length',
    'FOUTSPANNING_TYPE': 'Cable type',
    'N_OV_AFGEROND': 'Connections',
}


# Check for cable groups existing in the char excel and add characteristics to results:
for col in translate_columns.values():
    df_results['Characteristics', col] = pd.NA
counting = 0
missing = []
for i in df_results.index:
    group = str(df_results[('Scenario', 'Cable group')][i])
    if df_char['KABELGROEP'].str.contains(group).any():
        counting += 1
        index_nr = df_char.index[df_char['KABELGROEP'] == group].item()
        for col in translate_columns.keys():
            df_results.loc[i, ('Characteristics', translate_columns[col])] = df_char.at[index_nr, col]
    else:
        for col in translate_columns.values():
            df_results.loc[i, ('Characteristics', col)] = 'Unknown'
        missing.append(group)
print('Cable groups OK:', counting)
print('Cable groups missing in char:', list(np.unique(missing)))
df_results[('Characteristics', 'Urbanism')] = df_results[('Characteristics', 'Urbanism')].astype(str).replace(
    translate_urban, value=None)


# define violation percentages
aspects = ['Voltage', 'Loading', 'Fuse']
df_res_percentages = df_results[['Scenario', 'Characteristics']].copy()
for col in df_results[aspects].columns:
    df_res_percentages[col] = df_results[col] / df_results[('Scenario', 'Connections')] * 100
df_res_percentages[('Total', 'Cable group')] = df_results[aspects].sum(axis=1)
df_res_percentages[('Total', 'Violated')] = np.zeros(len(df_res_percentages))
df_res_percentages.loc[df_res_percentages[('Total', 'Cable group')] > 0, [('Total', 'Violated')]] = 1
df_res_percentages.drop([('Total', 'Cable group')], axis=1, inplace=True)


# --------------------- Numerical results ---------------------
# Define aggregation shortcuts:
def get_totals(df, lvl1_cols: list):
    return df.loc[:, df.columns.get_level_values(1).isin(['Total'] + lvl1_cols)]

def get_specifics(df, scenario_cols: list):
    scenario_data = df[
        [('Scenario', coll) for coll in scenario_cols] + list(df.columns[df.columns.get_level_values(0).isin(aspects)])]
    return scenario_data[scenario_data.columns[~scenario_data.columns.get_level_values(1).isin(['Total'])]]

def get_grouped_scenario(df, scenario_cols: list):
    return df.groupby(by=[('Scenario', elem) for elem in scenario_cols]).mean()

def get_grouped(df, type: str, scenario_cols: list):
    return df.groupby(by=[(type, elem) for elem in scenario_cols]).mean()

def get_scenario_is(df, scenario: str, item):
    return df.loc[df[('Scenario', scenario)] == item]


def get_is(df, type, scenario: str, item):
    return df.loc[df[(type, scenario)] == item]

def get_smaller_is(df, type: str, scenario: str, item):
    data = df.loc[df[(type, scenario)] != 'Unknown'].dropna()
    return data.loc[data[(type, scenario)].astype(int) <= item]

def get_larger(df, type: str, scenario: str, item):
    data = df.loc[df[(type, scenario)] != 'Unknown'].dropna()
    return data.loc[data[(type, scenario)].astype(int) > item]

def get_unknown(df, type: str, scenario: str):
    data = df.loc[df[(type, scenario)] == 'Unknown'].dropna()
    return data

def get_statistics(data):
    clear = data.loc[data[aspects][data[aspects] == 0].dropna(axis=0, how='any').index, aspects + ['Scenario']].groupby(
        ('Scenario', 'Cable group')).mean()
    grouped = data.groupby(('Scenario', 'Cable group')).mean()
    tests = int(len(grouped))
    violations = int(len(grouped) - len(clear))
    ratio = round(violations / tests * 100, 2)
    return [tests, ratio]


# Print numerical results: scenarios
nr_results = pd.DataFrame(index=['Cable group tests', 'violated [%]'])
for lamp in ['LED', 'old']:
    nr_results[lamp] = get_statistics(get_scenario_is(df_results, 'Light type', lamp))
for P in list(np.unique(df_results[('Scenario', 'Charger power')])):
    nr_results[P] = get_statistics(get_scenario_is(df_results, 'Charger power', P))
nr_results['Total'] = get_statistics(df_results)
nr_results['Total (no 75)'] = get_statistics(df_results.loc[df_results[('Scenario', 'Charger power')] != 75])
nr_results = nr_results.T
print(nr_results)
# nr_results.to_clipboard()

# Print numerical results: characteristics
cp = [0.92, 2.3, 3.7]
headers = ['Cable group tests', '{}kW: violated [%]'.format(cp[0]),
           '{}kW: violated [%]'.format(cp[1]), '{}kW: violated [%]'.format(cp[2])]
charger_results = pd.DataFrame(index=headers)
res_1_data = get_scenario_is(df_results, 'Charger power', cp[0])
res_2_data = get_scenario_is(df_results, 'Charger power', cp[1])
res_3_data = get_scenario_is(df_results, 'Charger power', cp[2])
def get_all(part, char, lim):
    if part == 'smaller':
        a = get_statistics(get_smaller_is(res_1_data, 'Characteristics', char, lim))
        b = [get_statistics(get_smaller_is(res_2_data, 'Characteristics', char, lim))[1]]
        c = [get_statistics(get_smaller_is(res_3_data, 'Characteristics', char, lim))[1]]
    if part == 'larger':
        a = get_statistics(get_larger(res_1_data, 'Characteristics', char, lim))
        b = [get_statistics(get_larger(res_2_data, 'Characteristics', char, lim))[1]]
        c = [get_statistics(get_larger(res_3_data, 'Characteristics', char, lim))[1]]
    if part == 'is':
        a = get_statistics(get_is(res_1_data, 'Characteristics', char, lim))
        b = [get_statistics(get_is(res_2_data, 'Characteristics', char, lim))[1]]
        c = [get_statistics(get_is(res_3_data, 'Characteristics', char, lim))[1]]
    if part == 'unknown':
        a = get_statistics(get_unknown(res_1_data, 'Characteristics', char))
        b = [get_statistics(get_unknown(res_2_data, 'Characteristics', char))[1]]
        c = [get_statistics(get_unknown(res_3_data, 'Characteristics', char))[1]]
    return a + b + c
if len(df_results) > 0:
    charger_results['until 300m'] = get_all('smaller', 'Length', 300)
    charger_results['> 300m'] = get_all('larger', 'Length', 300)
    charger_results['length unknown'] = get_all('unknown', 'Length', 300)
    charger_results['until 1960'] = get_all('smaller', 'Decade', 1960)
    charger_results['after 1960'] = get_all('larger', 'Decade', 1960)
    charger_results['age unknown'] = get_all('unknown', 'Decade', 1960)
    charger_results['SWAP cable'] = get_all('is', 'Cable type', 'SWAP')
    charger_results['Single cable'] = get_all('is', 'Cable type', 'single')
    charger_results['cable type unknown'] = get_all('unknown', 'Cable type', 'single')
    charger_results['until 30 lamps'] = get_all('smaller', 'Connections', 30)
    charger_results['> 30 lamps'] = get_all('larger', 'Connections', 30)
    charger_results['# lamps unknown'] = get_all('unknown', 'Connections', 30)
charger_results['Total'] = get_statistics(res_1_data) + [get_statistics(res_2_data)[1]] + [get_statistics(res_3_data)[1]]
charger_results = charger_results.T
print(charger_results)
# charger_results.to_clipboard()

df = df_results.loc[df_results[('Characteristics', 'Connections')] != 'Unknown'].copy()
a = []
for i in df.index:
    a.append(df.loc[i, ('Scenario', 'Connections')] - df.loc[i, ('Characteristics', 'Connections')])
df.loc[:, ('Characteristics', 'Diff')] = a
df[('Characteristics', 'Diff')].describe()

# --------------------- Graphs ---------------------

# Define data handling functions:
def get_pie_labels(labels, sizes):
    labels = [f'{l}, {s:0.1f}%' for l, s in zip(labels, sizes)]
    plt.legend(bbox_to_anchor=(0.85, 1), loc='upper left', labels=labels)
def add_line(ax, xpos, ypos):
    line = plt.Line2D([xpos, xpos], [ypos + .095, ypos],
                      transform=ax.transAxes, color='black')
    line.set_clip_on(False)  # [... + .1, ...]
    ax.add_line(line)
def label_len(my_index, level):
    labels = my_index.get_level_values(level)
    return [(k, sum(1 for i in g)) for k, g in groupby(labels)]
def label_group_bar_table(ax, fig, df_results):
    labels = ['' for item in ax.get_xticklabels()]
    ax.set_xticklabels(labels)
    ypos = -.1
    scale = 1. / df_results.index.size
    for level in range(df_results.index.nlevels)[::-1]:
        pos = 0
        for label, rpos in label_len(df_results.index, level):
            lxpos = (pos + .5 * rpos) * scale
            ax.text(lxpos, ypos, label, ha='center', transform=ax.transAxes)
            add_line(ax, pos * scale, ypos)
            pos += rpos
        add_line(ax, pos * scale, ypos)
        ypos -= .1
    ax.xaxis.set_label_coords(0.5, -0.25)
    fig.subplots_adjust(bottom=.1 * df_results.index.nlevels)
def get_grouped_char(df, char):
    df_char = df[[('Total', 'Violated'), ('Scenario', 'Cable group'), ('Characteristics', char)]
    ].groupby(by=[('Scenario', 'Cable group'),
                  ('Characteristics', char)]).sum().reset_index()  # any violation per cable group
    df_char.columns = ['Group', char, 'Violated']
    df_char[char] = df_char[char].astype(int, errors='ignore')
    df_char['count'] = np.zeros(len(df_char))
    for i in df_char.index:
        group = df_char.at[i, 'Group']
        count = len(df.loc[df[('Scenario', 'Cable group')] == group])
        df_char.loc[i, 'count'] = count
    df_char['% violated'] = round(df_char['Violated'] / df_char['count'] * 100, 2)
    df_2 = df_char.groupby(char)['% violated'].mean()
    df_2 = df_2.loc[df_2.index != 'Unknown'].astype(int)
    return df_2

# Define colors
colors_specific = ['#E76F51', '#EC8C74', '#1B655C', '#2A9D8F', '#4FD0C1', '#E9C46A', '#F4A261']  # 2, 3, 1
colors_total = [colors_specific[0], colors_specific[3], colors_specific[5]]  # 3

# Plot pie chart % of general and specific types of violations
scenario_cols = ['Light type', 'Charger power']
totals = get_grouped_scenario(get_totals(df_res_percentages, scenario_cols), scenario_cols)
specifics = get_grouped_scenario(get_specifics(df_res_percentages, scenario_cols), scenario_cols)
fig1, (ax1, ax2) = plt.subplots(1, 2, figsize=(8, 4))
ax1.pie(totals.sum(axis=0), autopct='%1.1f%%', pctdistance=0.85, colors=colors_total)
ax1.set_title('General violation type occurances')
ax1.legend(list(totals.columns.get_level_values(0)))
ax2.pie(specifics.sum(axis=0), autopct='%1.1f%%', pctdistance=0.85, colors=colors_specific)
ax2.set_title('Location of violation type occurances')
ax2.legend(list(specifics.columns.get_level_values(1)), loc='upper left')
plt.show()


# plot bar plot mean # of specific violations types per scenario type
scenario_cols = ['Light type', 'Charger power']
fig2, ax4 = plt.subplots(1, 1, figsize=(8, 5))
data = get_grouped_scenario(get_specifics(df_res_percentages, scenario_cols), scenario_cols)
data.set_axis(data.columns.map(', '.join), axis=1, inplace=False).plot(kind='bar', stacked=False, ax=ax4, color=colors_specific)
ax4.set_title('Average frequency of violation types per scenario')
label_group_bar_table(ax4, fig2, data)
ax4.set_xlabel('Light type, Charger power [kW]')
ax4.set_ylabel('Average % of locations in a cable group causing violation')
ax4.legend(loc='upper left')
plt.show()


# plot bar plot mean # general violations types per network size and age
df_age = get_grouped_char(df_res_percentages, 'Decade').reset_index()
df_size = get_grouped_char(df_res_percentages, 'Length').reset_index()
df_lamps = get_grouped_char(df_res_percentages, 'Connections').reset_index()

fig3, [ax7, ax8] = plt.subplots(1, 2, figsize=(11.2, 4))
df_age['% violated'].plot(kind='bar', ax=ax7, color=colors_total[0])
ax7.set_title('Network age')
labels = ['{}'.format(int(x)) for x in df_age['Decade']]
ax7.set_xticklabels(labels, rotation=0)
ax7.set_xlabel('Decade of installation')
ax7.set_ylabel('Average % of tests causing violation')

df_size['% violated'].plot(kind='bar', ax=ax8, color=colors_total[1])  # , color=colors_total[0])
ax8.set_title('Network size')
labels = ['{}'.format(int(x)) for x in df_size['Length']]
ax8.set_xticklabels(labels, rotation=0)
ax8.set_xlabel('Length of network cable groups [m]')
plt.suptitle('Average % of cable group tests causing violations per network characteristic', fontsize='x-large')
plt.show()

fig4, ax9 = plt.subplots(figsize=(6, 4))
df_lamps['% violated'].plot(kind='bar', ax=ax9, color=colors_total[2])
ax9.set_title('Network lamp connections')
labels = ['{}'.format(int(x)) for x in df_lamps['Connections']]
ax9.set_xticklabels(labels, rotation=0)
ax9.set_xlabel('Number of lamp connections')
ax9.set_ylabel('Average % of tests causing violation')
plt.show()


# end = time.time()
# time = end - start
# print('time: ', round(time, 2), 'sec')
