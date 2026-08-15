import pandas as pd
import numpy as np

def merge_description(df_詳細價目表):
    
    Column_單位=find_Column_單位的欄位(df_詳細價目表)
    for i in range(3, Column_單位):
        col = f"Column_{i}"
        
        mask = df_詳細價目表[col].notna()
        df_詳細價目表.loc[mask, "Column_2"] = (
        df_詳細價目表.loc[mask, "Column_2"].fillna("").astype(str)
        + " "
        + df_詳細價目表.loc[mask, col].astype(str))
        
    return df_詳細價目表

def find_Column_單位的欄位(df_詳細價目表):
    Column_單位=0
    for i in range(3, df_詳細價目表.shape[1]):
        if(df_詳細價目表[f"Column_{i}"].count())>50:
            Column_單位=i
            break
    return Column_單位
        