import pandas as pd
import os


# filename = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Technical Evaluation/Grid/Collected Data/Enexis/Round 2/20220307_Results_SL_R2102-02_gG-Light.xlsx'
# df = pd.read_excel(filename, sheet_name='Steekproef', engine='openpyxl', header=0)
# df = df[['KABELGROEP', 'NETSTATION', 'NETSTATION_STEDELIJKHEID', 'NETWERKTYPE', 'KABELGROEP_AANLEGJAAR', 'KABELGROEP_DECENIUM',
#          'KABELGROEP_LENGTE', 'NODE_MAX_FOUTSPANNING', 'N_OV_AANSLUITING']]
# print(df.head())
#
# df.count()
# df.KABELGROEP.unique()

results_dir = 'C:/Users/Elise/Desktop/Gaia/'  # where folders with excel files are saves from the Gaia results
data_dir = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Technical Evaluation/Grid/Collected Data/Enexis/Round 2/'

list_files = os.listdir(data_dir)
gnf_file_names = [filename[-14:-4] for filename in list_files if filename.endswith('.gnf')]

for dir_name in gnf_file_names:
    path = os.path.join(results_dir, dir_name)
    os.makedirs(path)