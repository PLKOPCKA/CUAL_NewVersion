from os import getlogin

# both lists need to be same length
bands = [0.4, 0.68, 0.78, 0.87, 0.93, 0.97, 1.01]
cuts = [95, 95, 95, 95, 95, 95, 95]
win_user = getlogin()
max_solve = 7200 # max time solve for the model
min_gap = 0.4 # min gap value for the model
scenario_name = 'CUAL_PL_2023_W03_v2' # name of the Scenario outputs in the Database Table
model_name = 'CUAL' # name of the model to choose from the CUAL folder
resolve = False  # if gap reached in max_solve time is lower than defined min_gap then split band into half
ISC_path = 'Transportation Lane Products Results.csv' # path for the download directory
fico_results_archive = r"E:\06.SCG Models\Poland_Customer_Allocation\FICO_results_archive"

if __name__ == '__main__':
    print(win_user)
    print(fico_results_archive)

