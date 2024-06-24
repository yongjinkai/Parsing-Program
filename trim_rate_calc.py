import pandas as pd
import os
import xlwings as xw
cwd = os.getcwd()
os.makedirs(os.path.join(cwd,"Results"),exist_ok=True)

# def formatter(df):
#     df.columns = df.iloc[0]
#     df = df.iloc[1:].reset_index(drop=True)
#     df = df.rename(columns={"WaferID":"Wafer"})   
#     return df

def actual_trim_calc(row):
    return row["Trim Amount"]/(float(row[f"{target_} current"])-float(row[f"{target_} previous"]))

def tabulate_data(df=pd.DataFrame):
    tabulated_columns=['LotID', 'WaferID', 'Trim', 'Count', 'Ave Trim Amt (nm)', 'Median Trim Amt (nm)', 'Ave ∆F (Current Trim-Previous Trim)',
                        'Median ∆F (Current Trim-Previous Trim)', 'Ave Trim Rate (nm/MHz)', 'Median Trim Rate (nm/MHz)']
    list_of_wafers = list(df['WaferID'].unique())
    global trimnumber
    trimnumber = df['Trim'][0]
    # print(list_of_wafers)
    tabulated_df = pd.DataFrame(columns=tabulated_columns)
    tabulated_df["WaferID"] = [i for i in df['WaferID'].unique()]
    tabulated_df["LotID"] = df["Lot ID"][0]
    tabulated_df["Trim"] = df["Trim"][0]
    # print(tabulated_df)
    for idx,i in enumerate(list_of_wafers):
        temp_df = df[df["WaferID"] == i].reset_index(drop=True)
        tabulated_df.loc[idx,"Count"] = len(temp_df["WaferID"])
        tabulated_df.loc[idx,"Ave Trim Amt (nm)"] = float(f'{temp_df["Trim Amount"].mean():.3f}')
        tabulated_df.loc[idx,"Median Trim Amt (nm)"] = float(f'{temp_df["Trim Amount"].median():.3f}')
        tabulated_df.loc[idx,"Ave ∆F (Current Trim-Previous Trim)"] = float(f'{temp_df["∆F (Current Trim-Previous Trim)"].astype(float).mean():.3f}')
        tabulated_df.loc[idx, "Median ∆F (Current Trim-Previous Trim)"] = float(f'{temp_df["∆F (Current Trim-Previous Trim)"].astype(float).median():.3f}')
        tabulated_df.loc[idx, "Ave Trim Rate (nm/MHz)"] = float(f'{temp_df["Actual_Trim_Rate"].mean():.3f}')
        tabulated_df.loc[idx, "Median Trim Rate (nm/MHz)"] = float(f'{temp_df["Actual_Trim_Rate"].median():.3f}')
    # print(tabulated_df)
    return tabulated_df

    

def data_merger(ibe,prev,curr,coords,target):
    #combining all IBE into single file
    global target_
    target_ = target
    concat_df = pd.DataFrame()
    
    for i in os.listdir(ibe):
        if os.path.splitext(i)[1].lower() == '.ibe':
            ibe_df = pd.read_csv(os.path.join(ibe,i),skiprows=[0,1],delimiter='\t')
            ibe_df.insert(0,"Trim",os.path.basename(i).split('-')[1][0:5])
            ibe_df.insert(0,"WaferID",os.path.basename(i).split('_')[1])
            ibe_df.insert(0,"Lot ID",os.path.basename(i).split('_')[0])
            concat_df = pd.concat([ibe_df,concat_df],ignore_index=True,axis=0) 
    column_names_changed = {f"{concat_df.columns[3]}":"real_xc",f"{concat_df.columns[4]}":"real_yc",f"{concat_df.columns[5]}":"Trim Amount"}
    concat_df.rename(columns=column_names_changed,inplace=True)


    #merging IBE data with x,y coordinate file data
    coord_df = pd.read_csv(coords)
    required_cols = ['xcoord','ycoord','real_xc','real_yc']
    if not all(x in required_cols for x in coord_df.columns ):
        raise Exception("Please upload coord file with the required columns: 'xcoord','ycoord','real_xc','real_yc' ")
    merged_ibe_df = concat_df.merge(coord_df, on=['real_xc','real_yc'],how="left")
    merged_ibe_df['WaferID'] = merged_ibe_df['WaferID'].astype(int,errors='ignore')

    #check for NA values(meaning IBE files contains coords not found in map)
    if merged_ibe_df.iloc[:,-2:].isna().sum().sum():
        raise Exception("Coordinate Map excel does not contain all IBE coordinates")
        

    columns_to_merge = ["Device","WaferID","ShotIndex","X","Y",f"{target}"]
    columns_to_merge2 = ["WaferID","X","Y",f"{target}"]
    prev_df = pd.read_excel(prev).rename(columns={'Wafer':'WaferID','Wafer ID':'WaferID','Shot Index': 'ShotIndex'})
    curr_df = pd.read_excel(curr).rename(columns={'Wafer':'WaferID','Wafer ID':'WaferID','Shot Index': 'ShotIndex'})


    temp_df = merged_ibe_df.merge(prev_df[columns_to_merge], left_on=["WaferID","xcoord","ycoord"], right_on=["WaferID","X","Y"],how="inner")
    temp_df = temp_df.rename(columns={f"{target}":f"{target} previous"})
   

    temp_df = temp_df.merge(curr_df[columns_to_merge2], on=["X","Y","WaferID"],how="inner")
    temp_df = temp_df.rename(columns={f"{target}":f"{target} current"})

    temp_df["Actual_Trim_Rate"] = temp_df.apply(actual_trim_calc,axis=1)
    temp_df["∆F (Current Trim-Previous Trim)"] = temp_df[f"{target} current"].astype(float) - temp_df[f"{target} previous"].astype(float)
    # print(temp_df)
    raw_df = temp_df[['Lot ID','WaferID','Trim','Device','ShotIndex',f'{target} previous',f'{target} current','∆F (Current Trim-Previous Trim)','Actual_Trim_Rate','Trim Amount']]
    raw_df = raw_df.sort_values(by=['WaferID','ShotIndex','Device'])
    return raw_df

def df_to_excel(currentdf,raw_trim_rate_df:pd.DataFrame,tabulated_df:pd.DataFrame):
    foldername = os.path.dirname(currentdf)
    finalfilename = os.path.join(foldername,f"{raw_trim_rate_df['Lot ID'][0]} {trimnumber} Trim Rate.xlsx")
    with pd.ExcelWriter(finalfilename,engine='openpyxl') as writer:
        raw_trim_rate_df.to_excel(writer,sheet_name="Trim Rate Raw",index=False)
        tabulated_df.to_excel(writer,sheet_name="Tabulation",index=False)
    
    app = xw.App(visible=False)
    wb = app.books.open(finalfilename)
    for sheet in wb.sheets:
        sheet.autofit()
    wb.save()
    app.quit()
    print(f'Completed. Results saved to {finalfilename}')
def trim_rate_calc_main(ibe,prev,curr,coord,target):
    print('Generating Trim Rate Report..')
    raw_trim_rate = data_merger(ibe,prev,curr,coord,target)
    tabulated_df = tabulate_data(raw_trim_rate)
    df_to_excel(curr,raw_trim_rate,tabulated_df)
    
    


if __name__ == "__main__":   
    ibe= "C:/Users/60061632\Desktop\Jinkai edits\Parsing programs\Sample files\Trimming calculation files/AEX853 TRIM2 IBE files"
    coord = 'C:/Users/60061632\Desktop\Jinkai edits\Parsing programs\Sample files\Trimming calculation files\B40 NTO Coord Map V0.csv'
    prev='C:/Users/60061632\Desktop\Jinkai edits\Parsing programs\Sample files\Trimming calculation files\AEX853_TRIM1_40FR.xlsx'
    curr='C:/Users/60061632\Desktop\Jinkai edits\Parsing programs\Sample files\Trimming calculation files\AEX853_TRIM2_40FR.xlsx'
    trim_rate_calc_main(ibe,prev,curr,coord,"F_RBE")
