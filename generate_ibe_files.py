import pandas as pd
import os
from decimal import Decimal, ROUND_HALF_UP
import xlwings as xw
import numpy as np


def gen_ibe_files(coord,esttrim):
    print('Generating IBE files...')
    finalexcelname = os.path.basename(esttrim)
    lotid = finalexcelname.split('_')[0]
    global currenttrim
    currenttrim = int(finalexcelname.split('_')[1][-1])
    finalfoldername = lotid + ' TRIM' + str(currenttrim+1) + ' IBE files'
    global finalfolderpath
    finalfolderparent = os.path.dirname(esttrim)
    finalfolderpath = os.path.join(finalfolderparent,finalfoldername)
    os.makedirs(finalfolderpath,exist_ok=True)

    merged_df,esttrim_df= merge_df(coord,esttrim)
    generate_ibe(merged_df,currenttrim+1,lotid,esttrim_df)
    

def merge_df(coord,esttrim):
    est_trim_df = pd.read_excel(esttrim,sheet_name='Est Trim').dropna()
    
    rawdatadf = pd.read_excel(esttrim,sheet_name='Raw Data')
    rawdatadf = rawdatadf.rename(columns={'Wafer ID':'WaferID','Wafer':'WaferID'})
    rawdatadf = rawdatadf[['WaferID','F_LBE','F_RBE','X','Y']]
    

    coord_df = pd.read_csv(coord)
    list_of_cols = coord_df.columns
    required_cols = ['xcoord','ycoord','real_xc','real_yc']
    if not all(x in required_cols for x in list_of_cols ):
        raise Exception("Please upload coord file with the required columns: 'xcoord','ycoord','real_xc','real_yc' ")

    
    merged_df = rawdatadf.merge(coord_df,how='inner',left_on=['X','Y'],right_on=['xcoord','ycoord'])
    merged_df = merged_df.drop(columns=['X','Y','xcoord','ycoord'])
    merged_df['Final Trim Amount (nm)'] = np.nan  
    merged_df = merged_df.merge(est_trim_df,how='inner',on='WaferID')
    merged_df = merged_df.sort_values(by='WaferID').reset_index(drop=True)
    

    for wafer in est_trim_df['WaferID'].unique():
        merged_mask = merged_df['WaferID'] == wafer
        est_trim_mask = est_trim_df['WaferID'] == wafer
        trim_target_selection = est_trim_df.loc[est_trim_mask,'F_LBE/F_RBE'].iloc[0]
        target_freq = est_trim_df.loc[est_trim_mask,'Target F (MHz)'].iloc[0]
        est_trim_rate = est_trim_df.loc[est_trim_mask,'Est. Trim Rate (nm/MHz)'].iloc[0]
        remaining_trim_amt = est_trim_df.loc[est_trim_mask,'Remaining Trim Amt (nm)'].iloc[0]
        merged_df.loc[merged_mask,'Final Trim Amount (nm)'] =  est_trim_rate*(target_freq - merged_df[merged_mask][trim_target_selection])-remaining_trim_amt
        rawdatatrimsites = len(rawdatadf[rawdatadf['WaferID']==wafer])
        ibefiletrimsites = len(merged_df[merged_mask])
        if rawdatatrimsites != ibefiletrimsites:
            print(f'Note: Wafer {wafer} has {rawdatatrimsites} lines of data but {ibefiletrimsites} generated trim sites')
    return merged_df,est_trim_df

def gen_summary_df(raw_ibe_df:pd.DataFrame,esttrimdf:pd.DataFrame) -> pd.DataFrame:
    
    
    agg_funcs = {
        'Trim Amount Max (nm)':'max',
        'Trim Amount Min (nm)':'min',
        'Trim Amount Median (nm)':'median',
        'Trim Amount Average (nm)': 'mean',
        'Trim Amount Range (nm)': lambda x: x.max()-x.min(),
        'Valid Trim Site (Trim amt >1nm) Count': lambda x: f'{(x>1).sum()}/{x.count()}',
        'Valid Trim Site (Trim amt >0.5nm) Count': lambda x: f'{(x>0.5).sum()}/{x.count()}',
    }

    df = raw_ibe_df.groupby('WaferID')["Final Trim Amount (nm)"].agg(**agg_funcs)
    df.insert(0,'Trim',currenttrim+1)
    
    # def left_or_right(row):
    #     if row['F_LBE/F_RBE'] == 'F_RBE':
    #         return row['Median F_RBE']
    #     elif row['F_LBE/F_RBE'] == 'F_LBE':
    #         return row['Median F_LBE']

    # esttrimdf.insert(2,'Median F_LBE/RBE', esttrimdf.apply(left_or_right,axis=1))
    # esttrimdf = esttrimdf.drop(columns=['Median F_LBE','Median F_RBE'])
    finaldf=esttrimdf.merge(df,how='inner',on='WaferID').round(3)

    return finaldf



def generate_ibe(df:pd.DataFrame,trimnum,lotid,esttrim_df):
    temp = df.loc[:,["WaferID","real_xc","real_yc","Final Trim Amount (nm)"]]
    temp = temp.sort_values(by=['real_xc','real_yc'],ascending=False).sort_values(by='WaferID').reset_index(drop=True)
    finaldf = gen_summary_df(temp,esttrim_df)
    rawibedata = temp.copy()

    # temp = temp.rename(columns={"real_xc":"%","real_yc":"Mo","Final Trim Amount (nm)":"10"})
    # print("before round:",temp["Mo"])
    temp["%"] = temp["real_xc"].apply(lambda x: Decimal(str(x)).quantize(Decimal('0.001'), rounding=ROUND_HALF_UP))
    temp["Mo"]=temp["real_yc"].apply(lambda x: Decimal(str(x)).quantize(Decimal('0.001'), rounding=ROUND_HALF_UP))
    temp["10"]=temp["Final Trim Amount (nm)"].map('{:.2f}'.format).astype(str)
    # print("after round:",temp["Mo"])
    columns = {"%": ["%","%mm"], "Mo":["y","mm"],"10":["removal","nm"]}
    headers = pd.DataFrame(columns)

    #generating individual IBE files
    for i in temp["WaferID"].unique():
        mask = temp["WaferID"] == i
        data = temp[mask][["%","Mo",'10']]
        final = pd.concat([headers,data])
        waferid = "0" + str(i) if i in range(1,10) else str(i)
        filename = str(lotid) + "_" + waferid + "_WAT-TRIM" + str(trimnum) + ".ibe"
        finalfilepath = os.path.join(finalfolderpath,filename)
        final.to_csv(finalfilepath,index=False,sep='\t')
    
    final_summary_filename = f'{lotid} Trim{trimnum} IBE summary.xlsx'
    final_summary_filepath = os.path.join(finalfolderpath,final_summary_filename)
    
    with pd.ExcelWriter(final_summary_filepath,engine='openpyxl') as writer:
        rawibedata.to_excel(writer,sheet_name='IBE files',index=False)
        finaldf.to_excel(writer,sheet_name='IBE Summary',index=False) 
        
    

    app = xw.App(visible=False)
    wb = app.books.open(final_summary_filepath)

    try:
        sheet = wb.sheets[1]
        
        table = sheet.range("A1").expand('table')
        sheet.api.ListObjects.Add(1,sheet.api.Range(table.address))

        leftcols = sheet.range('A1:H1')
        rightcols = sheet.range('I1:Q1')
        leftcols.color = (83,141,213)
        rightcols.color = (112,173,70)

        widths = [10,12,15,15,12,18,16,17,14,6,15,15,17,18,16,23,23]

        for idx,column in enumerate(table.columns):
            column.column_width = widths[idx]

        sheet.range('A:R').api.WrapText=True

        sheet2 = wb.sheets[0]
        sheet2.autofit()

    except Exception as e:
        print(f'Error: {e} encountered while formatting summary excel')

    wb.save(final_summary_filepath)
    wb.close()
    app.quit() 


    print(f'Completed. Results saved to {finalfolderpath}')

if __name__ == '__main__':
    for app in xw.apps:
        app.quit()