import os
import pandas as pd
import xlwings as xw
import numpy as np
from function import merge_description,find_Column_單位的欄位
import time

folder="詳細價目表"
file_directory = os.listdir(folder)
filepath = os.path.join(folder, file_directory[0])
header_row=int(input("輸入標頭列數: "))
df_詳細價目表 = pd.read_excel(filepath,header=None)

#--------------紀錄header----------------#
df_header = pd.read_excel(filepath,header=None,nrows=header_row)#header:dataframe
header_strig=str(df_header.iloc[0,0].strip())
print("刪除標頭:","\n",df_header)
#--------------紀錄header----------------#

#-------------刪除header-------------------#
rows_to_delete = []
for i in df_詳細價目表.index:
    value = str(df_詳細價目表.iloc[i, 0]).strip()
    if value.startswith(header_strig[0]):
        rows_to_delete.extend(range(i, i + header_row))     
df_詳細價目表 = (df_詳細價目表.drop(rows_to_delete, errors="ignore").reset_index(drop=True))
#-------------刪除header-------------------#

#-------------命名coloum(i)----------------#
df_詳細價目表.columns = [f'Column_{i}' for i in range(1, df_詳細價目表.shape[1] + 1)]#定義欄位名稱
df_詳細價目表.to_excel("output-詳細價目表0.xlsx", index=False)
#-------------命名coloum(i)----------------#

#df.shape[0]看列數
#df.shape[1]看欄數

#------------------合併多餘欄-----------------------#
Column_單位的欄位=find_Column_單位的欄位(df_詳細價目表)
Column_編碼的欄位=Column_單位的欄位+4 #編碼欄位固定為單位欄位後4欄
delete_coloum=[]
delete_coloum.extend(range(3, Column_單位的欄位))
delete_coloum.extend(range(Column_編碼的欄位,int(df_詳細價目表.shape[1])))
print(delete_coloum)

df_詳細價目表=merge_description(df_詳細價目表)
for i in range(Column_編碼的欄位+1,df_詳細價目表.shape[1]):
        col = f"Column_{i}"
        
        mask = df_詳細價目表[col].notna()
        df_詳細價目表.loc[mask, f'Column_{Column_編碼的欄位}'] = (
        df_詳細價目表.loc[mask, f'Column_{Column_編碼的欄位}'].fillna("").astype(str)
        + " "
        + df_詳細價目表.loc[mask, col].astype(str))
        
#-----------------合併多餘欄-----------------------#

#----------------刪除並重新命名欄-------------------#       
cols_to_drop = [f"Column_{i}" for i in delete_coloum]
df_詳細價目表.drop(columns=cols_to_drop,inplace=True)

df_詳細價目表.columns = [
    "項次",
    "項目及說明",
    "單位",
    "數量",
    "單價",
    "複價",
    "編碼(備註)"
]
df_詳細價目表.to_excel("output-詳細價目表1.xlsx", index=False)
#----------------刪除並重新命名欄-----------------#   

##--------------------合併列-------------------------##
new_rows = []
i = 0
總row數 = len(df_詳細價目表)
while i < 總row數:
    row = df_詳細價目表.iloc[i].copy()
    項次有值 = pd.notna(row["項次"])
    單位有值 = pd.notna(row["單位"])
    # ==========================
    # 條件一：章節列
    # ==========================
    if 項次有值 and not 單位有值:
        new_rows.append(row)
        i=i+1
        continue
    # ==========================
    # 條件二：主列
    # ==========================
    if 項次有值 and 單位有值:
        j=i+1
        while j < 總row數:
            next_row = df_詳細價目表.iloc[j]
            # 遇到下一個項次(大項或主列)停止
            if pd.notna(next_row["項次"]):
                break
            # ---------- 合併項目及說明 ----------
            if pd.notna(next_row["項目及說明"]):
                text = str(next_row["項目及說明"]).strip()
                if text != "":
                    row["項目及說明"] +=text
            # ---------- 合併編碼 ----------
            if pd.notna(next_row["編碼(備註)"]):
                code = str(next_row["編碼(備註)"]).strip()
                if code != "":
                    if pd.isna(row["編碼(備註)"]):
                        row["編碼(備註)"] = code
                    else:
                        row["編碼(備註)"] +=code
            j=j+1
        new_rows.append(row)
        i=j
        continue
    # ==========================
    # 其他列(例如空白列)
    # ==========================
    i += 1
##--------------------合併列-------------------------##

df_詳細價目表 = pd.DataFrame(new_rows).reset_index(drop=True)
df_詳細價目表.to_excel("output-詳細價目表2.xlsx", index=False)
input("Enter..結束")