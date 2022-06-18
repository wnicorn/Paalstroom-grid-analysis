import pandas as pd
import math
import numpy as np


filename = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Tech'\
           'nical Evaluation/Grid/Collected Data/Enexis/Round 2/20220307_Results_SL_R2102-02_gG-Light.xlsx'
df = pd.read_excel(filename, sheet_name='Steekproef', engine='openpyxl', header=0)
df = df[['KABELGROEP', 'NETSTATION_VESTIGING', 'NETSTATION', 'NETSTATION_STEDELIJKHEID', 'NETWERKTYPE', 'KABELGROEP_AANLEGJAAR',
         'KABELGROEP_DECENIUM', 'KABELGROEP_LENGTE', 'NODE_MAX_FOUTSPANNING', 'N_OV_AANSLUITING']]

df['KABELGROEP_LENGTE_TYPE'] = df['KABELGROEP_LENGTE']
df['FOUTSPANNING_TYPE'] = df['NODE_MAX_FOUTSPANNING']
df['N_OV_AFGEROND'] = df['N_OV_AANSLUITING']

for i in df.index:
    df['KABELGROEP_LENGTE_TYPE'].loc[i] = math.ceil(df['KABELGROEP_LENGTE'].loc[i].astype(float) / 100) * 100
    df['FOUTSPANNING_TYPE'].loc[i] = math.floor(df['NODE_MAX_FOUTSPANNING'].loc[i].astype(float) / 50) * 50
    df['N_OV_AFGEROND'].loc[i] = math.floor(df['N_OV_AANSLUITING'].loc[i].astype(float) / 5) * 5

print(df.head())

counts = {'total': len(df),
          'kabelgroep': len(df.KABELGROEP.unique()),
          'station': len(df.NETSTATION.unique()),
          'stedelijkheid': np.unique(df.NETSTATION_STEDELIJKHEID, return_counts=True),
          'Vestiging': np.unique(df.NETSTATION_VESTIGING, return_counts=True),
          'type': np.unique(df.NETWERKTYPE.dropna(inplace=False), return_counts=True),
          'aanleg_decenium': np.unique(df.KABELGROEP_DECENIUM, return_counts=True),
          'rond_lengte': np.unique(df.KABELGROEP_LENGTE_TYPE, return_counts=True),
          'foutspanning_type': np.unique(df.FOUTSPANNING_TYPE, return_counts=True),
          'n_ov': np.unique(df.N_OV_AFGEROND, return_counts=True)
          }
for entry in counts.keys():
    print(entry, ':')
    print(counts[entry], '\n')
# print(counts)
