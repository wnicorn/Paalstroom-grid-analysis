"""
Author: Elise van Wijngaarden
Updated: 20 June 2022

This script creates the folders that are needed in order to save the Gaia simulation results in the correct places.

INSTRUCTIONS:
only change the variables under 'Input':
- Edit 'results_dir' to the folder where all result folders should be saved (where Gaia will later save excel results).
- Edit 'data_dir' to the location where all the Gaia network files are located (on which the analyses will be performed)
"""

# ------------------- Input: -------------------------
results_dir = 'C:/Users/Elise/Desktop/Gaia/'  # where folders with excel files are saves from the Gaia results
data_dir = 'G:/.shortcut-targets-by-id/19-JqZkBCPZFbYYHLjE-rEmM8zRKqqge7/5LEF0 - SIP - Paalstroom/Research/2 Tech' \
           'nical Evaluation/Grid/Collected Data/Enexis/Round 2/additional'
# ---------------------------------------------------


import os

print('Creating folders...')
list_files = os.listdir(data_dir)
gnf_file_names = [filename[-14:-4] for filename in list_files if filename.endswith('.gnf')]

for dir_name in gnf_file_names:
    path = os.path.join(results_dir, dir_name)
    os.makedirs(path, exist_ok=True)

print('Done!')
