#!/usr/bin/env python
# coding: utf-8

# # Import necessary modules

print("Code Initializing...\n")

import os
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3' 

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    from tqdm import tqdm
except:
    install('tqdm')
    from tqdm import tqdm

print("Importing necessary librairies")

with tqdm(total=11) as pbar:

    try:
        import pandas as pd
    except:
        install('pandas')
        import pandas as pd
    pbar.update(1)

    try:
        import tensorflow
    except:
        install('tensorflow')
        import tensorflow
    pbar.update(1)
    
    try:
        import keras
    except:
        install('keras')
        import keras
    pbar.update(1)
    
    try:
        import numpy as np
    except:
        install('numpy')
        import pandas as np
    pbar.update(1)
    
    try:
        import pickle
    except:
        install('pickle')
        import pickle
    pbar.update(1)

    try:
        import tensorflow_hub as hub
    except:
        install('tensorflow-hub')
        import tensorflow_hub as hub
    pbar.update(1)

    try:
        import xlwings as xw
    except:
        install('xlwings')
        import xlwings as xw
    pbar.update(1)

    try:
        import openpyxl
    except:
        install('openpyxl')
        import openpyxl
    pbar.update(1)

    from tensorflow.keras.models import load_model
    pbar.update(1)
    import sys
    pbar.update(1)


# # Process the data

# ### Text embedding and encoding

    try:
        embedding = "https://tfhub.dev/google/nnlm-en-dim50/2"
    except:
        embedding = "./nnlm"
    hub_layer = hub.KerasLayer(embedding, input_shape=[], dtype=tensorflow.string, trainable=True)
    pbar.update(1)

def encode_string(string):
    return list(hub_layer([string]).numpy()[0])


# # Data Classification

def get_class_predicted(categories):
    return [np.argmax(category) for category in categories]


def row_is_empty(df, i):
    if df.iloc[i]["contact_job_function_selfrep"] != "":
        return False
    if df.iloc[i]["contact_job_role"] != "":
        return False
    if df.iloc[i]["contact_job_title"] != "":
        return False
    return True

def read_excel(path, sheet_name):     
    buffer = StringIO()            
    Xlsx2csv(path, outputencoding="utf-8").convert(buffer,sheetname=sheet_name)          
    buffer.seek(0)    
    df = pd.read_csv(buffer, low_memory=False)    
    return df

# # Write the function that will do everything

def classify_people_jobs(link_to_dataset, link_to_model, link_to_classes):
    last_3_char_of_model_link = link_to_model[-3:]
    if last_3_char_of_model_link != ".h5":
        link_to_model += ".h5"
    
    model = load_model(link_to_model)

    print("\nLoading the dataset")
    if link_to_dataset[-3:] == "csv":
        df = pd.read_csv(link_to_dataset, sep = ',', engine='python')
    else:
        df = pd.read_excel(link_to_dataset, sheet_name="result")
    
    
    #replace NaN by ""
    df = df.replace(float("nan"), "")

    print("Dataset Loaded\n")
    
    with open(link_to_classes, "rb") as fp:
        classes = pickle.load(fp)
    

    print("Step 1/3 - Processing and encoding the data")
    X = []
    indices_to_fill = []
    for i in tqdm(range(len(df))):
        if df.iloc[i]["contact_job_function"] == "" and not row_is_empty(df, i):
            encode_roles = [encode_string(df.iloc[i]["contact_job_function_selfrep"])]
            encode_roles += [encode_string(df.iloc[i]["contact_job_role"])]
            encode_roles += [encode_string(df.iloc[i]["contact_job_title"])]
            X.append(encode_roles)
            indices_to_fill.append(i)
            
    X = np.array(X)
    try:
        nsamples, nx, ny = X.shape
        X = X.reshape((nsamples,nx*ny))
    
        print("\nStep 2/3 - Processing the classification of contact_job_function")
        all_class_prediction = model.predict(X, verbose=1)
        prediction = get_class_predicted(all_class_prediction)
    
        classes_predicted = [classes[x] for x in prediction]
    
        # ignore warnings for better clarity 
        import warnings
        warnings.filterwarnings('ignore')
    
        job_functions = np.array(df["contact_job_function"])
        for i in range(len(X)):
            index = indices_to_fill[i]
            job_functions[index] = classes_predicted[i]
        
        return job_functions
    except:
        print("\nData was already classified")
        return np.array(df["contact_job_function"])


### Apply our functions to get results ###

link_to_dataset = sys.argv[1]
link_to_model = sys.argv[2]
link_to_classes = sys.argv[3]

job_functions = classify_people_jobs(link_to_dataset, link_to_model, link_to_classes)

### Use xlwings to write the results onto our excel workbook ###

wb = xw.Book(link_to_dataset)

sheet = wb.sheets['result']

col_contact_job_function = 1
while sheet.range((1, col_contact_job_function)).value != "contact_job_function" and sheet.range((1, col_contact_job_function)).value != "":
    col_contact_job_function += 1


#fill the column "contact_job_function"
print("\nStep 3/3 - Writing the data in the excel workbook")
for i in tqdm(range(len(job_functions))):
    sheet.range((i+2, col_contact_job_function)).value = job_functions[i]

print("Done")