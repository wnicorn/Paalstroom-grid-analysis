import pandas as pd

# prep
filename = 'EDITED_nederweert-sl-ndw.laurs.xlsx'
engine = 'openpyxl'
sheets = ['Knooppunten', 'Takken', 'Elementen', 'Hoofdkabels']
nodes = ['UL1', 'UL2', 'UL3']
loading = 'Belastinggraad'
connections = ['UL1N', 'UL2N', 'UL3N']
sort = 'Soort'
trafo = 'transformator'
cable = 'kabel'
voltages = nodes + connections
include = voltages + [loading, sort]

voltage_max = 253
voltage_min = 207

voltage_results = {}
loading_results = {}
voltage_violations = {}
loading_violations = {}
for elem in sheets:
    voltage_results[elem] = {}
    loading_results[elem] = []
    voltage_violations[elem] = 0
    loading_violations[elem] = 0
loading_results[trafo] = []
loading_violations[trafo] = 0


# --- data analysis ---
df = {}
df_v_results = {}
# Load files
for elem in sheets:
    # import files as dataframes per sheet
    df[elem] = pd.read_excel(filename, engine=engine, sheet_name=elem, header=0, skiprows=[1])
    # only include relevant columns
    df[elem] = df[elem][[name for name in include if name in df[elem].columns]]
    # drop columns that contain NaN values because they are not valuable
    df[elem] = df[elem].dropna()

    # check voltages
    if any(voltages in df[elem].columns.to_list() for voltages in voltages):
        # check per phase voltage:
        for voltage in [col for col in voltages if col in df[elem].columns]:
            voltage_results[elem][voltage] = []
            for i in df[elem].index:
                if df[elem][voltage][i] > voltage_max or df[elem][voltage][i] < voltage_min:
                    voltage_results[elem][voltage].append(1)
                else:
                    voltage_results[elem][voltage].append(0)
        # aggregate phases to sum violation at # of nodes:
        df_v_results[elem] = pd.DataFrame(voltage_results[elem])
        for i in df_v_results[elem].index:
            # print(df_v_results[elem].iloc[i])
            if df_v_results[elem].iloc[i].sum() > 0:
                voltage_violations[elem] += 1
    # check loadings
    if loading in df[elem].columns:
        for i in df[elem].index:
            if sort in df[elem].columns:
                if df[elem].at[i, sort] == cable:
                    if df[elem][loading][i] > 100:
                        loading_results[elem].append(1)
                    else:
                        loading_results[elem].append(0)
                elif df[elem][sort].loc[i] == trafo:
                    if df[elem][loading][i] > 100:
                        loading_results[trafo].append(1)
                    else:
                        loading_results[trafo].append(0)
            else:
                if df[elem][loading][i] > 100:
                    loading_results[elem].append(1)
                else:
                    loading_results[elem].append(0)
    # sum number of violations
    loading_violations[elem] = sum(loading_results[elem])
    loading_violations[trafo] = sum(loading_results[trafo])

total_voltage_violations = sum(voltage_violations.values())
total_loading_violations = sum(loading_violations.values())

# print results
print('voltage violations: ', voltage_violations)
print('loading violations: ', loading_violations)
print('total voltage violations: ', total_voltage_violations)
print('total loading violations: ', total_loading_violations)

# --- Graphics ---
