import tkinter as tk
from tkinter import filedialog,messagebox,ttk,font
from final import generate
import os,time,json
import pandas as pd
from trim_rate_calc import trim_rate_calc_main
from generate_ibe_files import gen_ibe_files
import threading
PADY = 5 
CWD = os.getcwd()

def extr_data(row=0):
    def rawdatabutton():
        filepath = filedialog.askdirectory(title="Select raw data folder")
        if filepath:
            extr_data_entry.delete(0,'end')
            extr_data_entry.insert(0,f'{filepath}')
    extr_data_button = ttk.Button(text = 'S2P file path',command=rawdatabutton,width=22)
    extr_data_button.grid(column=0,row=row,padx=10,pady=PADY)
    global extr_data_entry
    extr_data_entry = ttk.Entry(width = 70)
    extr_data_entry.grid(column=1,row=row,padx=10,pady=PADY,columnspan=5,sticky='w')
    # extr_data_entry.insert(0,entries['rawdata'])
    return extr_data_entry
def device_list(row=1):
    def devicelistbutton():
        filepath = filedialog.askopenfilename(title="Select Device List")
        if filepath:
            devicelist_entry.delete(0,'end')
            devicelist_entry.insert(0,f'{filepath}')
    devicelist_button = ttk.Button(text='Device List',command = devicelistbutton,width=22)
    devicelist_button.grid(column=0,row=row,padx=10,pady=PADY)
    global devicelist_entry
    devicelist_entry = ttk.Entry(width = 70)
    devicelist_entry.grid(column=1,row=row,padx=10,pady=PADY,columnspan=5,sticky='w')
    # devicelist_entry.insert(0,entries['devicelist'])
def mapping_table(row=2):
    def mappingtablebutton():
        filepath = filedialog.askopenfilename(title="Select Mapping Table")
        if filepath:
            mappingtable_entry.delete(0,'end')
            mappingtable_entry.insert(0,f'{filepath}')
    mappingtable_button = ttk.Button(text = 'Mapping Table',command = mappingtablebutton,width=22)
    mappingtable_button.grid(column=0,row=row,padx=10,pady=PADY)
    global mappingtable_entry
    mappingtable_entry = ttk.Entry(width = 70)
    mappingtable_entry.grid(column=1,row=row,padx=10,pady=PADY,columnspan=5,sticky='w')
    # mappingtable_entry.insert(0,entries['mappingtable'])
def s2p_key(row=3):
    s2pkey_label = ttk.Label(text='S2P key: \n eg. RA15')
    s2pkey_label.grid(column=0,row=row)
    global s2pkey_entry
    s2pkey_entry = ttk.Entry()
    s2pkey_entry.grid(column=1,row=row,padx=10,pady=PADY,sticky='w')
    # s2pkey_entry.insert(0,entries['testmap'])
def dropdown(row=3):
    global parseinfodf
    try:
        parseinfodf = extract_parse_info(parse_info_path)
        global device_options
        device_options = list(parseinfodf.columns[1:])
        parseinfoexist = True
    except FileNotFoundError:
        device_options = []
        parseinfoexist = False

    dropdown_label = ttk.Label(text='Select device:')
    dropdown_label.grid(row=row,column=3,padx=10,pady=PADY)
    global selected_option
    selected_option = tk.StringVar()
    selected_option.set('Select Device')

    
    if parseinfoexist:
        selected_option.set('Select Device')
        # if entries['dropdown_selection'] and entries['dropdown_selection'] in device_options:
        #     selected_option.set(entries['dropdown_selection'])
        #     on_option_select(entries['dropdown_selection'])  
        dropdown_menu =  tk.OptionMenu(window,selected_option, *device_options,command=on_option_select)
    else:
        dropdown_menu =  tk.OptionMenu(window,selected_option,None)
    dropdown_menu.grid(row=row,column=4,sticky='w')
    
    def show_error():
        error_message = f'Input fields manually or upload "Parse Information.xlsx" into {CWD}/Data Files'
        messagebox.showerror("Error", error_message)

    def on_click(event):
        nonlocal parseinfoexist
        if not parseinfoexist:
            try:
                extract_parse_info(parse_info_path)
                parseinfoexist=True
                print(' "Parse Information.xlsx" found! Click again')
            except FileNotFoundError:
                print('Input fields manually or upload "Parse Information.xlsx" into "Data Files" and try again')
                show_error()
            dropdown_menu.destroy()
            dropdown(row=row)

    dropdown_menu.bind("<Button-1>", on_click)
def on_option_select(selection):
    for index,entry in enumerate(entrylist):
        if index == 12:
            continue
        entry.delete(0,tk.END)
        paramvalue = parseinfodf[selection][index]
        # print(paramvalue,index)
        if type(paramvalue) == float and paramvalue.is_integer():
            entry.insert(0,int(paramvalue))
        else:
            entry.insert(0,paramvalue)
def searchmethod(row=5):
    search_method_label = ttk.Label(text='Freq Search Method:')
    search_method_label.grid(column=0,row=row,padx=10,pady=PADY,sticky='w')
    global selected_method
    selected_method = tk.StringVar(value="Outwards")
    search_method_frame = ttk.Frame()
    inwards_radiobutton = ttk.Radiobutton(search_method_frame, text="Inwards", variable=selected_method, value="Inwards")
    outwards_radiobutton = ttk.Radiobutton(search_method_frame, text="Outwards(default)", variable=selected_method, value="Outwards")
    outwards_radiobutton.pack(side=tk.LEFT)
    inwards_radiobutton.pack(side=tk.LEFT)
    search_method_frame.grid(column=1,row=row,padx=10,pady=PADY)
def sep_horizontal(row=4):
    separator_horizontal = ttk.Separator(orient='horizontal')
    separator_horizontal.grid(row=row, column=0, columnspan=5, sticky='ew', padx=5, pady=PADY)
def freq_start(row=6):
    freqstart_label = ttk.Label(text='Freq Start(MHz): ')
    freqstart_label.grid(column=0,row=row,padx=10,pady=PADY,sticky='w')
    global freqstart_entry
    freqstart_entry = ttk.Entry()
    freqstart_entry.grid(column=1,row=row,padx=10,pady=PADY,sticky='w')
    # freqstart_entry.insert(0,entries['start'])
def freq_stop(row=7):
    freqstop_label = ttk.Label(text='Freq Stop(MHz): ')
    freqstop_label.grid(column=0,row=row,padx=10,pady=PADY,sticky='w')
    global freqstop_entry
    freqstop_entry = ttk.Entry()
    freqstop_entry.grid(column=1,row=row,padx=10,pady=PADY,sticky='w')
    # freqstop_entry.insert(0,entries['stop'])
def passband_start(row=8):
    passbandstart_label = ttk.Label(text='Passband 通带 Start(MHz): ')
    passbandstart_label.grid(column=0,row=row,padx=10,pady=PADY,sticky='w')
    global passbandstart_entry
    passbandstart_entry = ttk.Entry()
    passbandstart_entry.grid(column=1,row=row,padx=10,pady=PADY,sticky='w')
    # passbandstart_entry.insert(0,entries['stop'])
def passband_stop(row=9):
    passbandstop_label = ttk.Label(text='Passband 通带 Stop(MHz): ')
    passbandstop_label.grid(column=0,row=row,padx=10,pady=PADY,sticky='w')
    global passbandstop_entry
    passbandstop_entry = ttk.Entry()
    passbandstop_entry.grid(column=1,row=row,padx=10,pady=PADY,sticky='w')
    # passbandstop_entry.insert(0,entries['stop'])
def sep_vertical(row=5):
    separator_vertical = ttk.Separator(orient='vertical')
    separator_vertical.grid(row=row,column=2, rowspan=14, sticky='ns',padx=5,pady=PADY )
def il_freq_1(row=5):
    il_freq_1_label = ttk.Label(text='IL_Freq_1(MHz): ')
    il_freq_1_label.grid(row=row,column=3,sticky='w')
    global il_freq_1_entry
    il_freq_1_entry = ttk.Entry()
    il_freq_1_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def il_freq_2(row=6):
    il_freq_2_label = ttk.Label(text='IL_Freq_2(MHz): ')
    il_freq_2_label.grid(row=row,column=3,sticky='w')
    global il_freq_2_entry
    il_freq_2_entry = ttk.Entry()
    il_freq_2_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def rej_freq_1(row=7):
    rej_freq_1_label = ttk.Label(text='Rej_Freq_1(MHz): ')
    rej_freq_1_label.grid(row=row,column=3,sticky='w')
    global rej_freq_1_entry
    rej_freq_1_entry = ttk.Entry()
    rej_freq_1_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def rej_freq_2(row=8):
    rej_freq_2_label = ttk.Label(text='Rej_Freq_2(MHz): ')
    rej_freq_2_label.grid(row=row,column=3,sticky='w')
    global rej_freq_2_entry
    rej_freq_2_entry = ttk.Entry()
    rej_freq_2_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def rej_freq_3(row=9):
    rej_freq_3_label = ttk.Label(text='Rej_Freq_3(MHz): ')
    rej_freq_3_label.grid(row=row,column=3,sticky='w')
    global rej_freq_3_entry
    rej_freq_3_entry = ttk.Entry()
    rej_freq_3_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def bw_il_1(row=10):
    bw_il_1_label = ttk.Label(text='BW_IL_1(dB): ')
    bw_il_1_label.grid(row=row,column=3,sticky='w')
    global bw_il_1_entry
    bw_il_1_entry = ttk.Entry()
    bw_il_1_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def bw_il_2(row=11):
    bw_il_2_label = ttk.Label(text='BW_IL_2(dB): ')
    bw_il_2_label.grid(row=row,column=3,sticky='w')
    global bw_il_2_entry
    bw_il_2_entry = ttk.Entry()
    bw_il_2_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def bw_il_3(row=12):
    bw_il_3_label = ttk.Label(text='BW_IL_3(dB): ')
    bw_il_3_label.grid(row=row,column=3,sticky='w')
    global bw_il_3_entry
    bw_il_3_entry = ttk.Entry()
    bw_il_3_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def il_lbe(row=13):
    il_lbe_label = ttk.Label(text='IL_LBE(dB): ')
    il_lbe_label.grid(row=row,column=3,sticky='w')
    global il_lbe_entry
    il_lbe_entry = ttk.Entry()
    il_lbe_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')    
def il_rbe(row=14):
    il_rbe_label = ttk.Label(text='IL_RBE(dB): ')
    il_rbe_label.grid(row=row,column=3,sticky='w')
    global il_rbe_entry
    il_rbe_entry = ttk.Entry()
    il_rbe_entry.grid(row=row,column=4,padx=0,pady=PADY,sticky='w')
def sep_horizontal2(row=10):
    separator_horizontal2 = ttk.Separator(orient='horizontal')
    separator_horizontal2.grid(row=row, column=0, columnspan=2, sticky='ew', padx=5, pady=PADY)
def r_off_bw_label(row=11):
    r_off_bw_label = ttk.Label(text='Left Roff =Left Freq delta between BW1 and BW2\ni.e Roff_L = F_BW1_L - F_BW2_L ')
    r_off_bw_label.grid(row=row,column=0,columnspan=2,padx=10,sticky='nw',pady=0,rowspan=3)
def r_off_bw1(row=12):
    r_off_bw1_label = ttk.Label(text='BW1(dB): ')
    r_off_bw1_label.grid(row=row,column=0,sticky='w',pady=PADY,padx=10)
    global r_off_bw1_entry
    r_off_bw1_entry = ttk.Entry()
    r_off_bw1_entry.grid(row=row,column=1,pady=PADY,sticky='w',padx=10)
def r_off_bw2(row=13):
    r_off_bw2_label = ttk.Label(text='BW2(dB): ')
    r_off_bw2_label.grid(row=row,column=0,sticky='w',padx=10)
    global r_off_bw2_entry
    r_off_bw2_entry = ttk.Entry()
    r_off_bw2_entry.grid(row=row,column=1,padx=10,pady=PADY,sticky='w')
def calc(row=19):
    global calculate_button
    calculate_button = ttk.Button(text='Calculate', command=calculate)
    calculate_button.grid(column=0,columnspan=10,row=row,sticky='ew')
def trim_rate_calc_btn(row=20):
    def show_trim_calc_btn():
        if trimcalc_frame.winfo_ismapped():
            trimcalc_frame.grid_forget()
        else: 
            gen_ibe_frame.grid_forget()
            trimcalc_frame.grid(column=0,columnspan=100,rowspan=4,sticky='ew')
    trim_rate_calc_button = ttk.Button(text='Trim Rate Calculator ↓↓', command=show_trim_calc_btn)
    trim_rate_calc_button.grid(column=0,columnspan=2,row=row,sticky='ew')
def generate_IBE_btn(row=20):
    def show_gen_IBE_btn():
        if gen_ibe_frame.winfo_ismapped():
            gen_ibe_frame.grid_forget()
        else:
            trimcalc_frame.grid_forget()
            gen_ibe_frame.grid(column=0,columnspan=100,rowspan=4,sticky='ew')
    generate_IBE_button = ttk.Button(text='IBE Files Generator ↓↓', command=show_gen_IBE_btn)
    generate_IBE_button.grid(column=3,columnspan=2,row=row,sticky='ew')
def trimming_frame(row=21):
    global trimcalc_frame
    trimcalc_frame = tk.Frame()
    trimcalc_frame.grid(row=row,column=0,columnspan=100,rowspan=4,sticky='ew')
    trimcalc_frame.grid_forget()

    global gen_ibe_frame
    gen_ibe_frame = tk.Frame()
    gen_ibe_frame.grid(row=row,column=0,columnspan=100,rowspan=4,sticky='ew')
    gen_ibe_frame.grid_forget()  


    def trim_rate_entries(row=1):
        trim_target_label = ttk.Label(trimcalc_frame,text='Trim Target Selection: ',)
        trim_target_label.grid(column=0,row=0)
        global selected_target
        selected_target = tk.StringVar(value='F_RBE')
        selected_target_frame = tk.Frame(trimcalc_frame)
        selected_target_frame.grid(column=1,row=0)
        f_lbe_radiobutton = ttk.Radiobutton(selected_target_frame,text='F_LBE',variable=selected_target,value='F_LBE')
        f_lbe_radiobutton.pack(side=tk.LEFT)
        f_rbe_radiobutton = ttk.Radiobutton(selected_target_frame,text='F_RBE',variable=selected_target,value='F_RBE')
        f_rbe_radiobutton.pack(side=tk.LEFT) 

        def ibe_path_button_func():
            filepath = filedialog.askdirectory(title="Select IBE folder")
            if filepath:
                ibe_path_entry.delete(0,'end')
                ibe_path_entry.insert(0,f'{filepath}')
        ibe_path_button = ttk.Button(trimcalc_frame,text= "IBE Folder Path",width=20,command=ibe_path_button_func)
        ibe_path_button.grid(column=0, row=row,padx=5,pady=5)
        
        global ibe_path_entry
        ibe_path_entry = ttk.Entry(trimcalc_frame)
        ibe_path_entry.grid(column=1,row=row,columnspan=100,sticky='ew', pady=5)
        def coord_map_button_func():
            filepath = filedialog.askopenfilename(title="Select Coord Map")
            if filepath:
                coord_map_entry1.delete(0,'end')
                coord_map_entry1.insert(0,f'{filepath}')
        coord_map_button = ttk.Button(trimcalc_frame,text= "Coord Map",width=20,command=coord_map_button_func)
        coord_map_button.grid(column=0, row=row+1,padx=5,pady=5)
        
        global coord_map_entry1
        coord_map_entry1 = ttk.Entry(trimcalc_frame)
        coord_map_entry1.grid(column=1,row=row+1,columnspan=100,sticky='ew', pady=5)
        
        def prev_trim_button_func():
            filepath = filedialog.askopenfilename(title="Select Previous Trim final excel")
            if filepath:
                previous_trim_entry.delete(0,'end')
                previous_trim_entry.insert(0,f'{filepath}')
        previous_trim_button = ttk.Button(trimcalc_frame,text= "Prev trim final excel",width=20,command=prev_trim_button_func)
        previous_trim_button.grid(column=0,row=row+2,padx=5,pady=5)
        
        global previous_trim_entry
        previous_trim_entry = ttk.Entry(trimcalc_frame)
        previous_trim_entry.grid(column=1,row=row+2,columnspan=100,sticky='ew',pady=5)
        
        def curr_trim_button_func():
            filepath = filedialog.askopenfilename(title="Select Current Trim final excel")
            if filepath:
                current_trim_entry.delete(0,'end')
                current_trim_entry.insert(0,f'{filepath}')
        current_trim_button = ttk.Button(trimcalc_frame,text= "Curr trim final excel",width=20,command=curr_trim_button_func)
        current_trim_button.grid(column=0,row=row+3,padx=5,pady=5)
        
        global current_trim_entry
        current_trim_entry = ttk.Entry(trimcalc_frame)
        current_trim_entry.grid(column=1,row=row+3,columnspan=100,sticky='ew',pady=5)
        
        calc_trim_rate_button = ttk.Button(trimcalc_frame,text= "Calculate Trim Rate",width=100,command=calc_trim_rate)
        calc_trim_rate_button.grid(column=0,columnspan=100,sticky='ew',row=row+4,pady=5)
    
    def gen_ibe_entries(row=0):  
        def coordmapbtn():
            filepath = filedialog.askopenfilename(title="Select Coordinate Map")
            if filepath:
                coord_map_entry2.delete(0,'end')
                coord_map_entry2.insert(0,f'{filepath}')
        coord_map_button = ttk.Button(gen_ibe_frame,text= "Coord Map",width=20,command=coordmapbtn)
        coord_map_button.grid(column=0,row=row+1,padx=5,pady=5,sticky='ew')
        global coord_map_entry2
        coord_map_entry2 = ttk.Entry(gen_ibe_frame)
        coord_map_entry2.grid(column=1,row=row+1,columnspan=100,sticky='ew',pady=5,padx=5)
        
        def esttrimbtn():
            filepath = filedialog.askopenfilename(title="Select Final Results Excel")
            if filepath:
                finalresults_esttrim_entry.delete(0,'end')
                finalresults_esttrim_entry.insert(0,f'{filepath}')
        finalresults_esttrim_button = ttk.Button(gen_ibe_frame,text= "Final excel with filled est trim",command=esttrimbtn)
        finalresults_esttrim_button.grid(column=0,row=row+2,padx=5,pady=5)
        global finalresults_esttrim_entry
        finalresults_esttrim_entry = ttk.Entry(gen_ibe_frame)
        finalresults_esttrim_entry.grid(column=1,row=row+2,columnspan=100,sticky='ew',pady=5,padx=5)

        gen_ibe_file_button = ttk.Button(gen_ibe_frame,text= "Generate IBE Files",width=100,command=gen_ibe)
        gen_ibe_file_button.grid(column=0,columnspan=100,sticky='ew',row=row+3,pady=5)
    trim_rate_entries()
    gen_ibe_entries()
def calc_trim_rate():
    ibe = ibe_path_entry.get()
    prev = previous_trim_entry.get()
    curr = current_trim_entry.get()
    coords = coord_map_entry1.get()
    target = selected_target.get()
    save_json(entry_history_dict)
    trim_rate_calc_main(ibe,prev,curr,coords,target)
def gen_ibe():
    coord = coord_map_entry2.get()
    esttrim = finalresults_esttrim_entry.get()
    save_json(entry_history_dict)
    gen_ibe_files(coord,esttrim)

def save_json(entry_history_dict):
    data = {}
    with open(json_path,'w',encoding='utf-8') as file:
        for key,entry in entry_history_dict.items():
            if entry.get():
                data[key]=entry.get()   
        json.dump(data,file)

def calculate():
    
    rawdata = extr_data_entry.get()
    devicelist = devicelist_entry.get()
    mappingtable = mappingtable_entry.get()
    testmap = s2pkey_entry.get()
    start = freqstart_entry.get()
    stop = freqstop_entry.get()
    search_method = selected_method.get()
    passband_start_ = passbandstart_entry.get()
    passband_stop_ = passbandstop_entry.get()
    ilfreq1 = 0 if il_freq_1_entry.get() == '' else il_freq_1_entry.get()
    ilfreq2 = 0 if il_freq_2_entry.get() == '' else il_freq_2_entry.get()
    rejfreq1 = 0 if rej_freq_1_entry.get() == '' else rej_freq_1_entry.get()
    rejfreq2 = 0 if rej_freq_2_entry.get() == '' else rej_freq_2_entry.get()
    rejfreq3 = 0 if rej_freq_3_entry.get() == '' else rej_freq_3_entry.get()
    bwil1 = 0 if bw_il_1_entry.get() == '' else bw_il_1_entry.get()
    bwil2 = 0 if bw_il_2_entry.get() == '' else bw_il_2_entry.get()
    bwil3 = 0 if bw_il_3_entry.get() == '' else bw_il_3_entry.get()
    illbe = il_lbe_entry.get()
    ilrbe = il_rbe_entry.get()
    roff1 = r_off_bw1_entry.get()
    roff2 = r_off_bw2_entry.get()

    save_json(entry_history_dict)

    calculate_button.config(state=tk.DISABLED)
    
    def thread_target():
        start_time = time.time()
        generate(rawdata,mappingtable,devicelist,testmap,start,stop,passband_start_,passband_stop_,search_method,ilfreq1,ilfreq2,rejfreq1,rejfreq2,rejfreq3,bwil1,bwil2,bwil3,illbe,ilrbe,roff1,roff2)
        calculate_button.config(state=tk.ACTIVE)
        end_time=time.time()
        timetaken = end_time-start_time
        print(f'time taken: {int(timetaken//60)} minutes {int(timetaken%60)} seconds')
    thread = threading.Thread(target=thread_target)
    thread.start()

    
def extract_parse_info(filepath):
    parseinfodf = pd.read_excel(filepath, skiprows=2,nrows=15).fillna('')
    parseinfodf = parseinfodf.iloc[:,1:].astype(object)
    return parseinfodf
def main():
    extr_data()
    device_list()
    mapping_table()
    s2p_key()
    sep_horizontal()
    sep_vertical()
    freq_start()
    freq_stop()
    searchmethod()
    passband_start()
    passband_stop()
    calc()
    il_freq_1()
    il_freq_2()
    rej_freq_1()
    rej_freq_2()
    rej_freq_3()
    bw_il_1()
    bw_il_2()
    bw_il_3()
    il_lbe()
    il_rbe()
    sep_horizontal2()
    r_off_bw_label()
    r_off_bw1()
    r_off_bw2()
    trimming_frame()
    trim_rate_calc_btn()
    generate_IBE_btn()
    dropdown()
    global entrylist
    entrylist = [il_freq_1_entry,il_freq_2_entry,rej_freq_1_entry,rej_freq_2_entry,rej_freq_3_entry,bw_il_1_entry,bw_il_2_entry,bw_il_3_entry,il_lbe_entry,il_rbe_entry,freqstart_entry,freqstop_entry,None,passbandstart_entry,passbandstop_entry]

    labels = ['s2pfileloc','devicelist','mappingtable','s2pkey','selecteddevice','ibepath','coordmap1','previoustrim','currenttrim','coordmap2','finalresults_esttrim']
    entry_history_list = [extr_data_entry,devicelist_entry,mappingtable_entry,s2pkey_entry,
                     selected_option,ibe_path_entry,coord_map_entry1,previous_trim_entry,current_trim_entry,coord_map_entry2,finalresults_esttrim_entry]
    
    global entry_history_dict
    entry_history_dict = dict(zip(labels,entry_history_list))

    
    if os.path.exists(json_path):
        with open(json_path, 'r') as f:
            data = json.load(f)
        for key, value in data.items():
            if key in entry_history_dict and key != 'selecteddevice':
                entry_history_dict[key].insert(0, value)
            if 'selecteddevice' in data and data['selecteddevice'] in device_options:
                selected_option.set(data['selecteddevice'])
                on_option_select(data['selecteddevice'])  

    
  
window = tk.Tk()
window.title("Filter Parser v1.4b")
window.config(width=500,height=500,padx=30,pady=15)


DataFilesFolder = os.path.join(CWD,'Data files')
os.makedirs(DataFilesFolder,exist_ok=True)
log_path = os.path.join(DataFilesFolder,"EntryHistory.txt")
json_path = os.path.join(DataFilesFolder,'EntryHist.json')
parse_info_path = os.path.join(DataFilesFolder,'Parse Information.xlsx')


if __name__ == '__main__':
    main()
    
window.mainloop()