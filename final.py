import os,all_result,time
import calculations_FR as c
import pandas as pd
import xlwings as xw
from openpyxl.worksheet.datavalidation import DataValidation

dtypedict = {'WaferID':int,'X':int,'Y':int}


def stripcor(cor):
    if cor[1] == 'N':
        return int(cor[2:]) * -1
    else:
        return int(cor[2:])
def mappingdf(MappingTablePath,DeviceListPath):
    mappingtabledf = pd.read_excel(MappingTablePath,skiprows=1,header=None)
    col1,col2,col3,col4=[i for i in mappingtabledf.columns]
    renamed_cols = {col1 : 'ShotIndex', col2 : 'Device', col3: None, col4: 'Coordinates'}
    mappingtabledf.rename(columns=renamed_cols,inplace=True)
    mappingtabledf = mappingtabledf[['ShotIndex','Device','Coordinates']]
    
    devicelistdf = pd.read_excel(DeviceListPath,skiprows=1,header=None)
    devicelistdf.rename(columns={devicelistdf.columns[0]:'Device'},inplace=True)
    combined = devicelistdf.merge(mappingtabledf,how = 'inner',right_on='Device',left_on='Device' if 'Device' in devicelistdf else 'device')
    combined['ShotIndex'] = combined['ShotIndex'].apply(lambda x: int(x.split('_')[1]))
    return combined
def generate(ChosenFolder,MappingTable,DeviceList,TestMap,start_freq,stop_freq,passband_start,passband_stop,direction='outwards',
             ILfreq1=0,ILfreq2=0,REJfreq1=0,REJfreq2=0,REJfreq3=0,BWIL1=0,BWIL2=0,BWIL3=0,ILLBE=0,ILRBE=0,roff1=0,roff2=0):
    print('generating.. do not close window')
    dct = {
'Coordinates':[],'WaferID':[],'X':[],'Y':[],"Max_DB(S2_1)":[],"F_Max_DB(S2_1)":[],"F1_(Max_DB-3db)":[],"F2_(Max_DB-3db)":[],"F2-F1":[],"IL_1":[],
"IL_2":[],"REJ_1":[],"REJ_2":[],"REJ_3":[],f"BW1_{BWIL1}":[],f"BW2_{BWIL2}":[],f"BW3_{BWIL3}":[],"F_LBE":[],"F_RBE":[],
"S11_MaxDB":[],"S22_MaxDB":[],"Roff31_R":[],"Roff31_L":[],"F_BW1_L":[],"F_BW1_R":[],"F_BW2_L":[],"F_BW2_R":[],"F_BW3_L":[],"F_BW3_R":[],
}
    if MappingTable == '' and DeviceList == '':
        parseall = True    
    else:
        combined = mappingdf(MappingTable,DeviceList)
        list_of_coords = list(combined['Coordinates'])
        parseall = False

    #initialising required variables
    
    wafernos = [str(i) if i>=10 else '0'+str(i) for i in range(1,26)]


    def gen_all_paths():
        temp_dict ={}
        for i in os.listdir(ChosenFolder):   
            if i in wafernos:
                waferpath = os.path.join(ChosenFolder, i,TestMap)
                try:
                    waferpath = os.path.join(waferpath,list(os.walk(waferpath))[0][1][0])
                    temp_dict[i] = waferpath
                    valid_waferpath = waferpath
                except IndexError:
                    print(f'Test Map {TestMap} not found in wafer {i}, moving on..')
        s2pfilename = os.listdir(valid_waferpath)[0] #Sample s2p file name: FS4101C_AAN163.16_08_RA15_TRIM2_Zone00001_XP014_YP090.s2p
        finalfoldername = s2pfilename.split('_')[1] + '_' + s2pfilename.split('_')[4] + '_' + s2pfilename.split('_')[3]
        global finalfolderpath
        finalfolderpath = os.path.join(ChosenFolder,finalfoldername)
        os.makedirs(finalfolderpath,exist_ok=True)

        finalfilename = s2pfilename.split('_')[1] + '_' + s2pfilename.split('_')[4] + '_' + s2pfilename.split('_')[3] + '.xlsx'
        finalfilepath = os.path.join(ChosenFolder,finalfilename)   
        return temp_dict,finalfilepath
    
    waferpathdict,finalfilepath= gen_all_paths()

    try: #check if final file excel is currently open (to prevent parsing all the way and then encountering permission error at the end)
        os.rename(finalfilepath,finalfilepath)
    except FileNotFoundError:
        pass  


    for key,value in waferpathdict.items():
        waferno = key
        waferresultfolder = os.path.join(finalfolderpath,f' Wafer {waferno} results')
        os.makedirs(waferresultfolder,exist_ok=True)
        waferresultfile = os.path.join(waferresultfolder,f'{waferno}.xlsx')
        all_s2p = os.listdir(value)
        if 'TEST_END' in all_s2p:
            all_s2p.remove('TEST_END')
        len_all_s2p = len(all_s2p)
        for idx,i in enumerate(all_s2p):
            s2pfilepath = os.path.join(value,i)
            coords= i[-15:-4] # example coords = XP013_YN124

            def generate_dict():
                subdf = c.s2pfile_to_df(s2pfilepath)[start_freq:stop_freq]       

                xcor = coords.split('_')[0]
                ycor = coords.split('_')[1]
                x = stripcor(xcor)
                y= stripcor(ycor)
                dct['Coordinates'].append(coords)
                dct['WaferID'].append(waferno)
                dct['X'].append(x)
                dct['Y'].append(y)
                
                maxdb,freq_maxdb = c.calc_maxdb(subdf,passband_start,passband_stop)
                dct['Max_DB(S2_1)'].append(round(maxdb,3))
                dct['F_Max_DB(S2_1)'].append(freq_maxdb)

                l,r = c.calc_minusxdb(subdf,freq_maxdb,maxdb-3,direction)
                dct['F1_(Max_DB-3db)'].append(l)
                dct['F2_(Max_DB-3db)'].append(r)
                dct['F2-F1'].append(r-l)

                for key_,value_ in {"IL_1":ILfreq1,"IL_2":ILfreq2,"REJ_1":REJfreq1,"REJ_2":REJfreq2,"REJ_3":REJfreq3}.items():
                    if value_ in [0,'','0']:
                        dct[key_].append(0)
                    else:
                        dct[key_].append(round(subdf.loc[float(value_),'s21 db'],3))
      
                
                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL1,direction)
                dct[f'BW1_{BWIL1}'].append(round(r-l,3))

                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL2,direction)
                dct[f'BW2_{BWIL2}'].append(round(r-l,3))

                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL3,direction)
                dct[f'BW3_{BWIL3}'].append(round(r-l,3))
                
                l,_ = c.calc_minusxdb(subdf,freq_maxdb,ILLBE,direction)
                _,r = c.calc_minusxdb(subdf,freq_maxdb,ILRBE,direction)
                dct['F_LBE'].append(l)
                dct['F_RBE'].append(r)
        
                dct['S11_MaxDB'].append(round(c.calc_s_maxdb(subdf,'s11',passband_start,passband_stop),3))
                dct['S22_MaxDB'].append(round(c.calc_s_maxdb(subdf,'s22',passband_start,passband_stop),3))
                
                l1,r1 = c.calc_minusxdb(subdf,freq_maxdb,roff1,direction)
                l2,r2 = c.calc_minusxdb(subdf,freq_maxdb,roff2,direction)
                dct['Roff31_L'].append(round(abs(l2-l1),3))
                dct['Roff31_R'].append(round(abs(r2-r1),3))
                
                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL1,direction)
                dct['F_BW1_L'].append(round(l,3))
                dct['F_BW1_R'].append(round(r,3))
                
                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL2,direction)
                dct['F_BW2_L'].append(round(l,3))
                dct['F_BW2_R'].append(round(r,3))
                
                l,r = c.calc_minusxdb(subdf,freq_maxdb,BWIL3,direction)
                dct['F_BW3_L'].append(round(l,3))
                dct['F_BW3_R'].append(round(r,3))

            if parseall or coords in list_of_coords:
                generate_dict()
            
            percentdone = round((idx+1)/len_all_s2p * 100)
            if (idx+1)%4 == 0 or percentdone == 100:
                print(f'Calculating for wafer {waferno}..  {percentdone}%',end=' \r')
        tempdf = pd.DataFrame(dct)
        tempdf=tempdf[tempdf['WaferID']==waferno]
        tempdf.to_excel(waferresultfile,index=False)
        print('')
        
    def gen_final_df(): #function to merge dct items with mapping table & device list into a dataframe
        temp = pd.DataFrame(dct)
        finaldf = pd.merge(combined,temp,how='inner',on='Coordinates')
        finaldf = finaldf[['WaferID','Device','ShotIndex',"Max_DB(S2_1)","F_Max_DB(S2_1)","F1_(Max_DB-3db)","F2_(Max_DB-3db)","F2-F1",
                        "IL_1","IL_2","REJ_1","REJ_2","REJ_3",f"BW1_{BWIL1}",f"BW2_{BWIL2}",f"BW3_{BWIL3}","F_LBE","F_RBE","S11_MaxDB","S22_MaxDB",
                        "Roff31_R","Roff31_L","F_BW1_L","F_BW1_R","F_BW2_L","F_BW2_R","F_BW3_L","F_BW3_R",'X','Y']]
        finaldf=finaldf.sort_values(by=['WaferID','ShotIndex','Device'],ascending=True)
        finaldf=finaldf.astype(dtypedict,errors='ignore')
        return finaldf
    
    if parseall:
        finaldf = pd.DataFrame(dct)
        finaldf = finaldf.drop(columns=['X','Y']).astype({'WaferID':int},errors='ignore')
        finaldf.to_excel(finalfilepath,sheet_name='Raw Data',index=False)
        app = xw.App(visible=False)
        wb = app.books.open(finalfilepath)
        rawdatasheet = wb.sheets['Raw Data']
        rawdatasheet.api.ListObjects.Add(1, rawdatasheet.api.Range(rawdatasheet.range('A1').expand('table').address))
        rawdatasheet.autofit()
        wb.save(finalfilepath)
        wb.close()
        app.quit()
    else:
        finaldf = gen_final_df()
        summarytab = all_result.gen_all_result(finaldf,BWIL1,BWIL2,BWIL3) #generate summary tab
        tabulation_df  = all_result.gen_tabulation_table(finaldf)   #generate tabulation tab
        est_trim_df = all_result.est_trim(finaldf) #generate estimated trim tab
        data_validation = DataValidation(type="list", formula1='"F_LBE,F_RBE"', allow_blank=True) 
        cell_range = f'B2:B{len(est_trim_df)+1}'
        

        def formatter_openpyxl():
            with pd.ExcelWriter(finalfilepath,engine='openpyxl') as writer: #writing all the dataframes to excel sheets with openpyxl engine
                finaldf.to_excel(writer,sheet_name='Raw Data',index=False)
                summarytab.to_excel(writer,sheet_name='Summary')
                tabulation_df.to_excel(writer,sheet_name='Tabulation',index=False)
                est_trim_df.to_excel(writer,sheet_name='Est Trim',index=False)
                writer.sheets['Summary'].delete_rows(3)
                writer.sheets['Est Trim'].add_data_validation(data_validation)
                data_validation.add(cell_range) 
        
        def formatter_xlwings():
            app = xw.App(visible=False)
            wb = app.books.open(finalfilepath) #Further formatting using xlwings
            summarysheet = wb.sheets['Summary']
            summarysheet.range('A:A').api.Delete()
            rawdatasheet = wb.sheets['Raw Data']
            rawdatasheet.api.ListObjects.Add(1, rawdatasheet.api.Range(rawdatasheet.range('A1').expand('table').address))
            trimsheet = wb.sheets['Est Trim']
            for sheet in wb.sheets:
                sheet.autofit()

            for i in range(len(est_trim_df)):
                x= i+2
                trimsheet.range(f'F{x}').formula = f'=IF(B{x}="F_RBE",E{x}-D{x},IF(B{x}="F_LBE",E{x}-C{x},0))'
                trimsheet.range(f'H{x}').formula = f'=G{x}*F{x}'
            
            trimsheet.range('B1').api.Interior.Color = 65535
            trimsheet.range('E1').api.Interior.Color = 65535
            trimsheet.range('G1').api.Interior.Color = 65535
            trimsheet.range('I1').api.Interior.Color = 65535
            wb.save(finalfilepath)
            wb.close()
            app.quit()
        formatter_openpyxl()
        formatter_xlwings()

    print('Completed')
    print('Results saved to ',finalfilepath)
    return finalfilepath
if __name__ == "__main__":
    mappingtablepath = '../Sample files\Mapping Table/N41 NTO/N41_41FR_Mapping Table v3.xlsx'
    devicelistpath = '../Sample files\Device List/N41 NTO/Filter_Main.xlsx'
    samples2pfile = "../Sample files/7. FS4101N1_AEP319_TRIM5_41FR_0328 (filter)/02/41FR\TRIM5/FS4101N1_AEP319_02_41FR_TRIM5_Zone00082_XN012_YP015.s2p"
    df = mappingdf(mappingtablepath,devicelistpath)
