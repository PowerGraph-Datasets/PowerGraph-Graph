import openpyxl
import numpy as np
import pandas as pd

# Specify the path to your Excel file
models_path = 'C:\\Users\\avarbella\\OneDrive - ETH Zurich\\Documents\\01_GraphGym\\PowerGraph-master\\code\\model\\'
rnd_seeds = [0, 100, 300, 700, 1000]
tasks = ['regression', 'binary', 'multiclass']
powergrids = ['ieee24', 'ieee39', 'uk', 'ieee118']
models = ['gcn', 'gin', 'gat', 'transformer']
# Specify the sheet name
sheet_name = 'Metrics'
sheet_regression = {
    'Power grid': ['ieee24', '', '', '', 'ieee39', '', '', '', 'uk', '', '', '', 'ieee118', '', '', ''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer'],
    'mse': [],
    'rmse': [],
}

sheet_binary = {
    'Power grid': ['ieee24', '', '', '', 'ieee39', '', '', '', 'uk', '', '', '', 'ieee118', '', '', ''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer'],
    'accuracy': [],
    'f1score': [],
    'balanced_accuracy': [],
    'roc_auc': [],
    'precision': [],
    'recall': [],
}

sheet_multiclass = {
    'Power grid': ['ieee24', '', '', '', 'ieee39', '', '', '', 'uk', '', '', '', 'ieee118', '', '', ''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer', 'gcn', 'gin', 'gat', 'transformer'],
    'accuracy': [],
    'f1score': [],
    'balanced_accuracy': [],
    'roc_auc': [],
    'precision': [],
    'recall': [],
}


results = {}

for powergrid in powergrids:
    results[powergrid] = {}
    for model in models:
        results[powergrid][model] = {}
        for task in tasks:
            results[powergrid][model][task] = {}
            if task == 'regression':
                metrics = ['mse', 'rmse']
                results[powergrid][model][task][metrics[0]] = {}
                results[powergrid][model][task][metrics[1]] = {}
                mse = []
                rmsescore = []
            else:
                metrics = ['accuracy', 'f1score', 'balanced_accuracy', 'roc_auc', 'precision', 'recall']
                results[powergrid][model][task][metrics[0]] = {}
                results[powergrid][model][task][metrics[1]] = {}
                results[powergrid][model][task][metrics[2]] = {}
                results[powergrid][model][task][metrics[3]] = {}
                results[powergrid][model][task][metrics[4]] = {}
                results[powergrid][model][task][metrics[5]] = {}
                accuracy = []
                f1score = []
                balanced_accuracy = []
                roc_auc = []
                precision = []
                recall = []

            for rnd_seed in rnd_seeds:
                specific_excel_file_path = models_path + powergrid + '\\' + 'summary' + powergrid + '_' + model + '_' + task + '_3l_16h_' + str(rnd_seed) +'s' + '.xlsx'
                print(specific_excel_file_path)
                workbook = openpyxl.load_workbook(specific_excel_file_path)
                sheet = workbook[sheet_name]
                if task == 'regression':
                    mse.append(sheet['B2'].value)
                    rmsescore.append(sheet['B3'].value)
                else:
                    accuracy.append(sheet['B2'].value)
                    f1score.append(sheet['B3'].value)
                    balanced_accuracy.append(sheet['B4'].value)
                    roc_auc.append(sheet['B5'].value)
                    precision.append(sheet['B6'].value)
                    recall.append(sheet['B7'].value)
                workbook.close()
            if task == 'regression':
                sheet_regression[metrics[0]].append(str(np.format_float_scientific(np.mean(mse), precision=4))+'±'+str(np.format_float_scientific(np.std(mse), precision=4)))
                sheet_regression[metrics[1]].append(str(np.format_float_scientific(np.mean(rmsescore), precision=4))+'±'+str(np.format_float_scientific(np.std(rmsescore), precision=4)))

            elif task == 'binary':
                sheet_binary[metrics[0]].append(str(np.mean(accuracy).round(4))+'±'+str(np.std(accuracy).round(4)))
                sheet_binary[metrics[1]].append(str(np.mean(balanced_accuracy).round(4))+'±'+str(np.std(balanced_accuracy).round(4)))
                sheet_binary[metrics[2]].append(str(np.mean(f1score).round(4))+'±'+str(np.std(f1score).round(4)))
                sheet_binary[metrics[3]].append(str(np.mean(roc_auc).round(4))+'±'+str(np.std(roc_auc).round(4)))
                sheet_binary[metrics[4]].append(str(np.mean(precision).round(4))+'±'+str(np.std(precision).round(4)))
                sheet_binary[metrics[5]].append(str(np.mean(recall).round(4))+'±'+str(np.std(recall).round(4)))

            else:
                sheet_multiclass[metrics[0]].append(str(np.mean(accuracy).round(4))+'±'+str(np.std(accuracy).round(4)))
                sheet_multiclass[metrics[1]].append(str(np.mean(balanced_accuracy).round(4))+'±'+str(np.std(balanced_accuracy).round(4)))
                sheet_multiclass[metrics[2]].append(str(np.mean(f1score).round(4))+'±'+str(np.std(f1score).round(4)))
                sheet_multiclass[metrics[3]].append(str(np.mean(roc_auc).round(4))+'±'+str(np.std(roc_auc).round(4)))
                sheet_multiclass[metrics[4]].append(str(np.mean(precision).round(4))+'±'+str(np.std(precision).round(4)))
                sheet_multiclass[metrics[5]].append(str(np.mean(recall).round(4))+'±'+str(np.std(recall).round(4)))

df_sheet_Regression = pd.DataFrame(sheet_regression)
df_sheet_binary = pd.DataFrame(sheet_binary)
df_sheet_multiclass = pd.DataFrame(sheet_multiclass)

excel_file_path = 'processed_results.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
    # Write each DataFrame to a different sheet
    df_sheet_Regression.to_excel(writer, sheet_name='regression', index=False)
    df_sheet_binary.to_excel(writer, sheet_name='binary', index=False)
    df_sheet_multiclass.to_excel(writer, sheet_name='multiclass', index=False)
"""
# Load the Excel workbook
workbook = openpyxl.load_workbook(excel_file_path)

# Select the desired sheet
sheet = workbook[sheet_name]

# Specify the cell or range of cells you want to retrieve
cell_value = sheet['A1'].value  # Change 'A1' to your desired cell reference

# Or, if you want to iterate through a range of cells, for example, column A from row 1 to 10
column_values = [sheet[f'A{i}'].value for i in range(1, 11)]

# Print the results
print(f"Value in A1: {cell_value}")
print(f"Values in column A from row 1 to 10: {column_values}")
"""
# Close the workbook when done
workbook.close()
