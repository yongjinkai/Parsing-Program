import pandas as pd 
import cmath


def s2pfile_to_df(s2pfile):
    df = pd.read_csv(s2pfile,delim_whitespace=True,skiprows=1,header=None,index_col=0)
    df = df.rename_axis('freq')
    df.index = df.index/1000000
    df['s11'] = df[1] + df[2] * 1j
    df['s22'] = df[7] + df[8] * 1j
    df['s21'] = df[3] + df[4] * 1j
    df['s11 db'] =20 * df['s11'].abs().apply(lambda x: cmath.log10(x).real)
    df['s22 db'] =20 * df['s22'].abs().apply(lambda x: cmath.log10(x).real)
    df['s21 db'] =20 * df['s21'].abs().apply(lambda x: cmath.log10(x).real)
    df = df.drop(columns=[1,2,3,4,5,6,7,8,'s11','s22','s21'])
    return df

def calc_maxdb(subdf:pd.DataFrame,start,stop):
    #N41 NTO maxdb is between start/stop freq, not between passband start/stop
    # subdf = subdf[start:stop]
    maxdb = subdf['s21 db'].max()
    freq_maxdb = subdf[subdf['s21 db'] == maxdb].index[0]
    return maxdb,freq_maxdb

def calc_s_maxdb(df,s:str,start,stop):
    # B40 NTO s_maxdb is between passband start/stop
    df = df.loc[start:stop]
    if 's11' in s:
        return df['s11 db'].max()
    if 's22' in s:
        return df['s22 db'].max()

def calc_minusxdb(subdf:pd.DataFrame,freq_maxdb,x,direction):
    if x == 0 or x == '':
        return 0,0
    x = float(x)
    if direction == 'Inwards': #boolean to put into ascending=bool
        leftdfsort = True
        rightdfsort = False
    else:
        leftdfsort = False
        rightdfsort = True
    subdfleft = subdf[:freq_maxdb]['s21 db']-x
    subdfleft=subdfleft.sort_index(ascending=leftdfsort)
    prev_sign = 1 if subdfleft.iloc[0] > 0 else 0
    freq_left=freq_right=0
    for idx,i in enumerate(subdfleft):
        curr_sign = 1 if i>0 else 0
        if curr_sign != prev_sign:
            if abs(subdfleft.iloc[idx])< abs(subdfleft.iloc[idx-1]):
                freq_left = subdfleft.index[idx]
                break
            else:
                freq_left = subdfleft.index[idx-1]
                break
    subdfright = subdf[freq_maxdb:]['s21 db']-x
    subdfright = subdfright.sort_index(ascending=rightdfsort)
    prev_sign = 1 if subdfright.iloc[0] > 0 else 0
    for idx,i in enumerate(subdfright):
        curr_sign = 1 if i>0 else 0
        if curr_sign != prev_sign:
            if abs(subdfright.iloc[idx])< abs(subdfright.iloc[idx-1]):
                freq_right = subdfright.index[idx]
                break
            else:
                freq_right = subdfright.index[idx-1]
                break
    return freq_left,freq_right

if __name__ == "__main__":
    start = 2300
    stop = 2800
    stepsize = 0.1
    samples2pfile = "../Sample files/7. FS4101N1_AEP319_TRIM5_41FR_0328 (filter)/02/41FR\TRIM5/FS4101N1_AEP319_02_41FR_TRIM5_Zone00082_XN012_YP015.s2p"
    sampledf = s2pfile_to_df(samples2pfile)
    samplesubdf = sampledf.loc[start:stop]
    
    maxdb,freq_maxdb = calc_maxdb(samplesubdf,start,stop)
    # print(f'{freq_maxdb=}')

    # freq_left,freq_right = calc_minusxdb(samplesubdf,freq_maxdb,-35)
    # print(freq_left,freq_right)

    # left2,right2=calc_minusxdbv2(samplesubdf,freq_maxdb,stepsize,-35.0)
    # print(left2,right2)
    # l3,r3=calc_minusxdb(samplesubdf,freq_maxdb,-35.0)
    # print(l3,r3)

