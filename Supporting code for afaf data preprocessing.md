
import glob
import subprocess

# doc to docx

for doc in glob.iglob("../*.doc"):
    subprocess.call(['soffice', '--headless', '--convert-to', 'docx', doc])
    
## Note: for rft to docx --> change .doc to rtf
    
# showing the docx file

glob.glob("*.docx")

# make files in folder and read the files

types = ("*.docx") # the tuple of file types
files_grabbed = []
for files in types:
    files_grabbed.extend(glob.iglob(files))
    
files_grabbed

# Process the Arabic document. 
# Example 1 

def EWP_AR(files_grabbed,path):    
    import pandas as pd
    from docx.api import Document
    import pandas as pd
    import re
    import docx2txt
    import os

    my_text = docx2txt.process(files_grabbed)
    my_text = re.sub(r'(\n\n\n\n)(\[[\d]+:[\d]+:[\d]+.[\d]+\])(\n\n)','\n', my_text) # regex same format

    my_text = my_text.split('\n')
    df = pd.DataFrame(my_text)
    df = df.iloc[3:] 				# dataframe locte from row 3
    nan_value = float("NaN")
    df.replace("", nan_value, inplace=True)
    df.dropna(inplace=True)
    
# save the file as cvs from dataframe. (same file name)    
    if not os.path.exists(path):
        os.makedirs(path)    
    df.to_csv(path+'/'+filename+'.csv',sep=';',header=None,index=False)
    return df

# Example 2
def LT_AR(files_grabbed,path):    
    import pandas as pd
    from docx.api import Document
    import pandas as pd
    import re
    import docx2txt
    import os

    my_text = docx2txt.process(files_grabbed)
    my_text1 = my_text.replace('\n\n',' ')
    my_text3 = my_text1.replace('?', '?\n')
    my_text4 = my_text3.replace('.','.\n')
    my_text2 = my_text4.split('\n')
    df = pd.DataFrame(my_text2)
    nan_value = float("NaN")
    df.replace("", nan_value, inplace=True)
    df.dropna(inplace=True)
    
    
    if not os.path.exists(path):
        os.makedirs(path)    
    df.to_csv(path+'/'+filename+'.csv',sep=';',header=None,index=False)
    return df

# Save the files in specified folder
for filename in files_grabbed:
    print(filename)
    try:
        EWP_AR(filename,'name of the folder')
    except:
        print("File corrpted:",filename)
        
        
######## Combine two files into One #########
# Read all csv file 
import glob

types = [("*.csv")] # the tuple of file types
files_grabbed = []
for files in types:
    files_grabbed.extend(glob.iglob(files))
    
    
    
# combine each file and save in single file
import pandas as pd
import csv
df1 = pd.read_csv('NR.ep04.en.docx.csv', error_bad_lines=False)
df2 = pd.read_csv('NR.ep04.ar.docx.csv', error_bad_lines=False)

df3 = pd.concat([df2, df1], axis=1)
df3.to_csv('NR.ep04.csv',sep=';', index=False)





















        
