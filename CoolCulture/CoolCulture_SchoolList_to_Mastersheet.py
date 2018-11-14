import sys
import pandas as pd
import glob
from pandas import ExcelWriter
from pandas import ExcelFile

outputCols = ['FAMILY NAME','TYPE','Program','SCHOOL NAME','GW ID','Pass ID #','Type','ADULT 1 FIRST NAME','ADULT 1 LAST NAME','ADULT 2 FIRST NAME','ADULT 2 LAST NAME','CHILD FIRST NAME','CHILD LAST NAME',"CHILD'S CLASS",'Adult 1 Email Address',"Adult 2 Email Address","STAFF TITLE","Staff Email"]
finalOutput = pd.DataFrame(columns=outputCols)

for index, file in enumerate(glob.glob("CoolCulture/Schools/*")):
    schoolName = file.split('/')[2].replace('.xlsx','')
    print(schoolName)
        
    familyOutput = pd.DataFrame(columns=outputCols)
    input = pd.read_excel(file, sheetname='Families', usecols=8)
    
    familyOutput["ADULT 1 FIRST NAME"] = input.iloc[:,0]
    familyOutput["ADULT 1 LAST NAME"] = input.iloc[:,1]
    familyOutput["ADULT 2 FIRST NAME"] = input.iloc[:,2]
    familyOutput["ADULT 2 LAST NAME"] = input.iloc[:,3]
    familyOutput['CHILD FIRST NAME'] = input.iloc[:,4]
    familyOutput['CHILD LAST NAME'] = input.iloc[:,5]
    familyOutput['CHILD\'S CLASS'] = input.iloc[:,6]
    familyOutput['Adult 1 Email Address'] = input.iloc[:,7]
    familyOutput['Adult 2 Email Address'] = input.iloc[:,8]
    familyOutput['Type'] = 'F'
    
    input = pd.read_excel(file, sheetname='Staff', usecols=5)
    
    staffOutput = pd.DataFrame(columns=outputCols)
    
    staffOutput["ADULT 1 FIRST NAME"] = input.iloc[:,0]
    staffOutput["ADULT 1 LAST NAME"] = input.iloc[:,1]
    staffOutput["ADULT 2 FIRST NAME"] = input.iloc[:,2]
    staffOutput["ADULT 2 LAST NAME"] = input.iloc[:,3]
    staffOutput["STAFF TITLE"] = input.iloc[:,4]
    staffOutput["Staff Email"] = input.iloc[:,5]
    staffOutput['Type'] = 'S'
    
    output = pd.DataFrame(columns=outputCols)
    output = pd.concat([familyOutput,staffOutput])
    
    output['TYPE'] = 'FAMILY'
    
    for i in range(output.shape[0]):
        name1 = output.iloc[i,:]['ADULT 1 LAST NAME']
        name2 = output.iloc[i,:]['ADULT 2 LAST NAME']
        if pd.isnull(name1):
             familyName = name2
        elif pd.isnull(name2):
            familyName = name1
        else:
            familyName = name1 + '/' + name2
        print(familyName + "-" + str(i))
        output.iat[i,0] = familyName
        
    output['SCHOOL NAME'] = schoolName
    output.sort_values(by='FAMILY NAME')
    finalOutput = pd.concat([finalOutput,output])

finalOutput.to_excel('CoolCulture/output/Mastersheet.xlsx',index=False)