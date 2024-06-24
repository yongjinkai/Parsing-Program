import pandas as pd

def gen_all_result(df:pd.DataFrame,BWIL1,BWIL2,BWIL3):
    params_list = ['Max_DB(S2_1)','IL_1','IL_2','REJ_1','REJ_2','REJ_3',f'BW1_{BWIL1}',f'BW2_{BWIL2}',f'BW3_{BWIL3}','F_LBE','F_RBE','S11_MaxDB','S22_MaxDB','Roff31_L','Roff31_R']
    summary = df[['WaferID']].drop_duplicates().reset_index()[['WaferID']]
    for i in params_list:
        if i in df.columns:
            if i in ['BW1_0','BW2_0','BW3_0']:
                continue
            agg_dict = {f'Mean {i}': 'mean' , f'Median {i}':'median', f'std {i}':'std',f'Max {i}':'max',f'Min {i}':'min',f'Range {i}':lambda x: x.max()-x.min()}
            temp_df = df.groupby('WaferID',as_index=False)[i].agg(**agg_dict)
            summary = summary.merge(temp_df,how='inner',on='WaferID')
            summary = summary.round(3)

    outer_levels = [''] + [i.split(' ')[1] for i in summary.columns[1:]]



    multi_index = pd.MultiIndex.from_tuples(zip(outer_levels,summary.columns))
    summary.columns = multi_index
    return summary

def gen_tabulation_table(df:pd.DataFrame):

    tabulation_count = df.groupby('WaferID',as_index=False)['Device'].count().rename(columns={'Device':'Count'})
    tabulation_1_5db = df[(df['Max_DB(S2_1)'] < -1.5)].groupby('WaferID',as_index=False)['Device'].count().rename(columns={'Device':'<-1.5db'})
    tabulation_2db = df[(df['Max_DB(S2_1)'] < -2)].groupby('WaferID',as_index=False)['Device'].count().rename(columns={'Device':'<-2db'})
    tabulation_min = df.groupby('WaferID',as_index=False)[['Max_DB(S2_1)']].min().rename(columns={'Max_DB(S2_1)':'Min'})  

    merged_tabulation = tabulation_1_5db.merge(tabulation_2db,on='WaferID',how='outer').merge(tabulation_min,on='WaferID',how='outer').merge(tabulation_count,on='WaferID',how='outer').fillna(0)
    return merged_tabulation[['WaferID','<-1.5db','<-2db','Min','Count']].sort_values(by='WaferID')

def est_trim(df:pd.DataFrame):
    
    median_f_lbe = df.groupby('WaferID',as_index=False)[['F_LBE']].median()
    median_f_rbe = df.groupby('WaferID',as_index=False)[['F_RBE']].median()
    est_trim_df = median_f_lbe.merge(median_f_rbe,how='outer',on='WaferID')
    est_trim_df['F_LBE/F_RBE'] = None
    est_trim_df = est_trim_df[['WaferID','F_LBE/F_RBE','F_LBE','F_RBE']]
    est_trim_df.rename(columns={'F_LBE':'Median F_LBE','F_RBE':'Median F_RBE'},inplace=True)
    est_trim_df['Target F (MHz)'] = None
    est_trim_df['ΔF (Target-F_RBE/F_LBE)'] = None
    est_trim_df['Est. Trim Rate (nm/MHz)'] = None
    est_trim_df['Est. TTL Trim Amount (nm)'] = None
    est_trim_df['Remaining Trim Amt (nm)'] = None
    return est_trim_df