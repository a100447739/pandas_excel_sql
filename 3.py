import os
import pandas as pd
import xlwings as xw
import numpy as np
import class_機電物件

folder="單價分析表"
file_directory = os.listdir(folder)
filepath = os.path.join(folder, file_directory[0])
df_單價分析表 = pd.read_excel(filepath, header=None)

'''print(df_單價分析表.shape[1])'''
df_單價分析表.columns = [f'Column_{i}' for i in range(1, df_單價分析表.shape[1] + 1)]#定義欄位名稱

df_單價分析表[df_單價分析表.columns[0]] = df_單價分析表[df_單價分析表.columns[0]].shift(1, fill_value=None)

df_單價分析表.to_excel("output_單價分析表.xlsx", index=True)       

for i in range (0,df_單價分析表.shape[0]):                                               #調整位置
    if not (pd.isna(df_單價分析表.loc[i, "Column_1"])):
        df_單價分析表.loc[i , "Column_2"] = df_單價分析表.loc[i-1, "Column_2"] 
        df_單價分析表.loc[i-1, "Column_2"] = None                 
        df_單價分析表.loc[i + 1, "Column_2"] = df_單價分析表.loc[i + 1, "Column_2"] 
        
df_單價分析表.to_excel("output_單價分析表.xlsx", index=True)       

#刪除不必要列--------------------------------------
物件_index=[]
內容物列高_index=[]      
for i in range (0,df_單價分析表.shape[0]):
    if not (pd.isna(df_單價分析表.loc[i, "Column_1"])):
        物件_index.append(i)
        count_row=0
        while(df_單價分析表.loc[i+count_row, "Column_2"]!="合計"):
            count_row=count_row+1
        內容物列高_index.append(count_row+i)                          #####合計調整
delete_index = []
for i in range(len(物件_index)-1):
    #print(物件_index[i],"   ",內容物列高_index[i])
    start_index=內容物列高_index[i]
    end_index=物件_index[i+1]
    delete_index.extend(range(start_index, end_index))
    
df_單價分析表.drop(index=delete_index, inplace=True)    
df_單價分析表.drop(index=0, inplace=True)

print(物件_index)
#print(delete_index) 
#刪除不必要列--------------------------------------------------

#尋找編碼(備註)欄數--------------------------------------------------
series=df_單價分析表.loc[1].to_list()
print(series)
count_col,i=0,0
while(series[i]!="編碼(備註)"):
    count_col= count_col+1
    i=i+1 
print(count_col)   
#check編碼錯位:
if count_col+1 != len(df_單價分析表.columns):                                          #如果欄位不正,則調整
    編碼錯位欄index_row=[]                                                             #前一欄移到正確欄
    編碼錯位欄index_row=df_單價分析表.index[df_單價分析表.iloc[:,count_col-1].notna()]   #iloc[x:y,count_col];第x~y row,欄=count_col(index)||df_單價分析表.index[count_col].notna() ; count_col會被識別成字串
    for i in range (len(編碼錯位欄index_row)):
        #print(編碼錯位欄index_row[i]," ",df_單價分析表.loc[編碼錯位欄index_row[i]].iloc[count_col-1])
        column_name_編碼錯位欄 = df_單價分析表.columns[count_col-1]
        column_name_編碼欄     = df_單價分析表.columns[count_col]                                                                               
        df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼欄]=df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼錯位欄]
        df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼錯位欄]=None
#---------------------------
if count_col + 1 < len(df_單價分析表.columns):                                         #檢查是否有下一欄
    編碼錯位欄index_row=[]                                                             #下一欄移到正確欄
    編碼錯位欄index_row=df_單價分析表.index[df_單價分析表.iloc[:,count_col+1].notna()]   #iloc[x:y,count_col];第x~y row,欄=count_col(index)||df_單價分析表.index[count_col].notna() ; count_col會被識別成字串
    for i in range (len(編碼錯位欄index_row)):
        #print(編碼錯位欄index_row[i]," ",df_單價分析表.loc[編碼錯位欄index_row[i]].iloc[count_col-1])
        column_name_編碼錯位欄 = df_單價分析表.columns[count_col+1]
        column_name_編碼欄     = df_單價分析表.columns[count_col]                                                                                
        df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼欄]=df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼錯位欄]
        df_單價分析表.loc[編碼錯位欄index_row[i],column_name_編碼錯位欄]=None                                     


#刪除不必要欄--------------------------------------------------



 
#df_單價分析表.drop(columns=df_單價分析表.columns[7:count_col],inplace=True)
#df_單價分析表.drop(columns=df_單價分析表.columns[count_col:],inplace=True)    
#刪除不必要欄--------------------------------------------------
'''for i in range(len(物件_index)):
    df_單價分析表.loc[i , "Column_4"]=None
    df_單價分析表.loc[i , "Column_5"]=None
    df_單價分析表.loc[i , "Column_6"]=None
    df_單價分析表.loc[i , "Column_7"]=None
print(物件_index) '''
 
   

df_單價分析表.to_excel("output_單價分析表.xlsx", index=True)
'''df_單價分析表.columns = ["項 次", "項  目  及  說  明","單 位","數 量","單 價","複 價","編碼(備註)"]'''

