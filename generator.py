# %% import package
import numpy as np
import pandas as pd
import glob
import os
import shutil
# %% parameters
LabName = 'Block 1 Full Report [Total Pts: 20 Score] |1586594'
# %%
fileNames = glob.glob('*.csv')
for i, ele in enumerate(fileNames):
    print(i)
    if i == 0:
        data = pd.read_csv(ele)
        data_NeedGrad = data[data[LabName] == 'Needs Grading']
    else: 
        data = pd.read_csv(ele)
        data_NeedGrad = data_NeedGrad.append(data[data[LabName] == 'Needs Grading'])
sourcesPath = './Rubric.pdf'
firstName  = list(data_NeedGrad['First Name'])
lastName = list(data_NeedGrad['Last Name'])
netID = list(data_NeedGrad['Username'])
# %%
for i in range(len(firstName)):
    targetPath = './sheet/' + lastName[i] + '_' + firstName[i] + '_' + netID[i] + '.pdf'
    if not os.path.exists(targetPath):
        shutil.copy(sourcesPath, targetPath)
# %%
filePath = glob.glob('./sheet/*.pdf')
for i in range(len(filePath)):
    path = filePath[i]
    email = filePath[i].split('.')[-2].split('_')[-1] 
    inputpdf = PdfFileReader(open(filePath[i], "rb"))
    output = PdfFileWriter()
    output.addPage(inputpdf.getPage(0))
    with open('./sheet/' + email + '.pdf', "wb") as outputStream:
        output.write(outputStream)
