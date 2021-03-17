# %%
import numpy as np
import pandas as pd
import glob
import os
import shutil
import random
import PySimpleGUI as sg
# %% 
filename = glob.glob('*.xlsx')
start_index = 'S'; end_index = 'AN'
num_question = 6
def colNameToNum(name):
    pow = 1
    colNum = 0
    for letter in name[::-1]:
            colNum += (int(letter, 36) -9) * pow
            pow *= 26
    return colNum
start_index = colNameToNum(start_index) - 1
end_index = colNameToNum(end_index) - 1
# %%
data = pd.read_excel(filename[0], skiprows = 2)
data.columns = data.iloc[0]
data = data[1:]
header = data.columns
print('Start column name is \n{}'.format(header[start_index]))
print('End column name is \n{}'.format(header[end_index]))
questions = header[start_index:end_index+1]

for i, row in data.iterrows():
    index = random.sample(range(0, len(questions)), num_question)
    index.sort()
    Name = row['First Name'] + row['Last Name']
    layout = [[sg.Text('Student Name: ' + Name)]]
    for q in index:
        question = questions[q]
        score = data.at[i, question]
        layout.append([sg.Text(question)])
        layout.append([sg.InputText(default_text = str(score), key = question)])
    layout.append([sg.Button('Submit'), sg.Button('Exit')])
    window = sg.Window('LiuHong ME 310 grader', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        
        if event == 'Submit':
            for q in index:
                question = questions[q]
                data.at[i, question] = float(values[question])
            break
    window.close()
data.to_excel('result.xlsx')

   
        




# %%
