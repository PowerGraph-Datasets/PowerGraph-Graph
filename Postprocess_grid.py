import openpyxl
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

# Specify the path to your Excel file
models_path = 'C:\\Users\\avarbella\\OneDrive - ETH Zurich\\Documents\\01_GraphGym\\PowerGraph-master\\code\\model\\'
rnd_seeds = [0, 100, 300, 700, 1000]
hyperparameters = ['1l_8h','2l_8h','3l_8h','1l_16h','2l_16h','3l_16h','1l_32h','2l_32h','3l_32h']
tasks = ['binary', 'multiclass','regression']
powergrids = ['ieee118']
n_bus = 118
models = ['gcn', 'gin', 'gat', 'transformer']
# Specify the sheet name
sheet_name = 'Metrics'
sheet_name_d = 'Data'

sheet_binary = {
    'hyperparameters': ['1l_8h','','','','2l_8h','','','','3l_8h','','','','1l_16h','','','','2l_16h','','','','3l_16h','','','','1l_32h','','','','2l_32h','','','','3l_32h','','',''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer'],
    'accuracy': [],
    'f1score': [],
    'balanced_accuracy': [],
    'roc_auc': [],
    'precision': [],
    'recall': [],
}
sheet_multiclass = {
    'hyperparameters': ['1l_8h','','','','2l_8h','','','','3l_8h','','','','1l_16h','','','','2l_16h','','','','3l_16h','','','','1l_32h','','','','2l_32h','','','','3l_32h','','',''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer'],
    'accuracy': [],
    'f1score': [],
    'balanced_accuracy': [],
    'roc_auc': [],
    'precision': [],
    'recall': [],
}

sheet_regression = {
    'hyperparameters': ['1l_8h','','','','2l_8h','','','','3l_8h','','','','1l_16h','','','','2l_16h','','','','3l_16h','','','','1l_32h','','','','2l_32h','','','','3l_32h','','',''],
    'MPL type': ['gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer','gcn', 'gin', 'gat', 'transformer'],
    'mse': [],
    'rmse': [],
}
results_binary = {}
results_multiclass = {}
results_regression = {}

for task in tasks:
    for layer in hyperparameters:
        results_binary[layer] = {}
        results_multiclass[layer] = {}
        results_regression[layer] = {}
        powergrid = powergrids[0]
        for model in models:
            results_binary[layer][model] = {}
            results_multiclass[layer][model] = {}
            results_regression[layer][model] = {}
            if task == 'binary':
                metrics = ['accuracy', 'f1score', 'balanced_accuracy', 'roc_auc', 'precision', 'recall']
                results_binary[layer][model][metrics[0]] = {}
                results_binary[layer][model][metrics[1]] = {}
                results_binary[layer][model][metrics[2]] = {}
                results_binary[layer][model][metrics[3]] = {}
                results_binary[layer][model][metrics[4]] = {}
                results_binary[layer][model][metrics[5]] = {}
                accuracy = []
                f1score = []
                balanced_accuracy = []
                roc_auc = []
                precision = []
                recall = []

            elif task == 'multiclass':
                metrics = ['accuracy', 'f1score', 'balanced_accuracy', 'roc_auc', 'precision', 'recall']
                results_multiclass[layer][model][metrics[0]] = {}
                results_multiclass[layer][model][metrics[1]] = {}
                results_multiclass[layer][model][metrics[2]] = {}
                results_multiclass[layer][model][metrics[3]] = {}
                results_multiclass[layer][model][metrics[4]] = {}
                results_multiclass[layer][model][metrics[5]] = {}
                accuracy = []
                f1score = []
                balanced_accuracy = []
                roc_auc = []
                precision = []
                recall = []


            else:
                metrics = ['mse', 'rmse']
                results_regression[layer][model][metrics[0]] = {}
                results_regression[layer][model][metrics[1]] = {}
                mse = []
                rmsescore = []

            for rnd_seed in rnd_seeds:
                specific_excel_file_path = models_path + powergrid + '\\' + 'summary' + powergrid + '_' + model + '_' + task + '_'+layer+'_' + str(rnd_seed) +'s' + '.xlsx'
                print(specific_excel_file_path)
                workbook = openpyxl.load_workbook(specific_excel_file_path)
                sheet = workbook[sheet_name]
                if task == 'binary':
                    accuracy.append(sheet['B2'].value)
                    f1score.append(sheet['B3'].value)
                    balanced_accuracy.append(sheet['B4'].value)
                    roc_auc.append(sheet['B5'].value)
                    precision.append(sheet['B6'].value)
                    recall.append(sheet['B7'].value)

                elif task == 'multiclass':
                    accuracy.append(sheet['B2'].value)
                    f1score.append(sheet['B3'].value)
                    balanced_accuracy.append(sheet['B4'].value)
                    roc_auc.append(sheet['B5'].value)
                    precision.append(sheet['B6'].value)
                    recall.append(sheet['B7'].value)

                else:
                    mse_value = sheet['B2'].value
                    mse.append(mse_value)
                    rmsescore.append(sheet['B3'].value)
                workbook.close()

            if task == 'binary':
                sheet_binary[metrics[0]].append(str(np.mean(accuracy).round(4)) + '±' + str(np.std(accuracy).round(4)))
                sheet_binary[metrics[1]].append(
                    str(np.mean(balanced_accuracy).round(4)) + '±' + str(np.std(balanced_accuracy).round(4)))
                sheet_binary[metrics[2]].append(str(np.mean(f1score).round(4)) + '±' + str(np.std(f1score).round(4)))
                sheet_binary[metrics[3]].append(str(np.mean(roc_auc).round(4)) + '±' + str(np.std(roc_auc).round(4)))
                sheet_binary[metrics[4]].append(
                    str(np.mean(precision).round(4)) + '±' + str(np.std(precision).round(4)))
                sheet_binary[metrics[5]].append(str(np.mean(recall).round(4)) + '±' + str(np.std(recall).round(4)))
            elif task == 'multiclass':
                sheet_multiclass[metrics[0]].append(str(np.mean(accuracy).round(4)) + '±' + str(np.std(accuracy).round(4)))
                sheet_multiclass[metrics[1]].append(
                    str(np.mean(balanced_accuracy).round(4)) + '±' + str(np.std(balanced_accuracy).round(4)))
                sheet_multiclass[metrics[2]].append(str(np.mean(f1score).round(4)) + '±' + str(np.std(f1score).round(4)))
                sheet_multiclass[metrics[3]].append(str(np.mean(roc_auc).round(4)) + '±' + str(np.std(roc_auc).round(4)))
                sheet_multiclass[metrics[4]].append(
                    str(np.mean(precision).round(4)) + '±' + str(np.std(precision).round(4)))
                sheet_multiclass[metrics[5]].append(str(np.mean(recall).round(4)) + '±' + str(np.std(recall).round(4)))

            else:
                sheet_regression[metrics[0]].append(str(np.format_float_scientific(np.mean(mse), precision=4))+'±'+str(np.format_float_scientific(np.std(mse), precision=4)))
                sheet_regression[metrics[1]].append(str(np.format_float_scientific(np.mean(rmsescore), precision=4))+'±'+str(np.format_float_scientific(np.std(rmsescore), precision=4)))

    if task == 'binary':
        df_sheet_binary = pd.DataFrame(sheet_binary)
    elif task == 'multiclass':
        df_sheet_multiclass = pd.DataFrame(sheet_multiclass)
    else:
        df_sheet_regression = pd.DataFrame(sheet_regression)


excel_file_path = f'processed_results_{powergrid}_cascades.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
    # Write each DataFrame to a different sheet
    df_sheet_binary.to_excel(writer, sheet_name='binary', index=False)
    df_sheet_multiclass.to_excel(writer, sheet_name='multiclass', index=False)
    df_sheet_regression.to_excel(writer, sheet_name='regression', index=False)
    for task in tasks:
        worksheet = writer.sheets[task]

# Close the workbook when done
workbook.close()
