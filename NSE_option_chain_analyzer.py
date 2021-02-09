#!/usr/bin/python3

import requests
from typing import Union, Optional, List, Dict, Tuple, TextIO, Any
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, \
    Label, Entry, SOLID, RIDGE, N, S, E, W, LEFT, messagebox
from tkinter.ttk import Combobox, Button, Notebook
import pandas as pd
import tksheet
import nsepy.live as nse
from datetime import datetime
import time
from sys import exit

class NSE:
    def stop_all(self)->None:
        self.stop = True 
    def main_recursive(self)->None:
        if(self.first_run):
            #print("FIRST time")
            self.refresh_data()
            self.first_run = False
        self.curr_time = time.time()
        time_passed = int(self.curr_time-self.prev_time)
        if(time_passed>self.interval):
            self.refresh_data()
        else:
            self.sh_window.after((10 * 1000),self.main_recursive)
            return
        if(not self.stop):
            self.sh_window.after((10 * 1000),self.main_recursive)
            return
            
    def __init__(self,window: Tk)->None:
        self.first_run: bool = True
        self.dict_dfs_INDEX: dict[pd.DataFrame] = {}
        self.dict_dfs_STOCK: dict[pd.DataFrame] = {}
        self.nb_names: List[String] = ['CE-OTM','PE-OTM','CE-FAR_OTM','PE-FAR_OTM','LTP','PCR']
        for i in self.nb_names:
            self.dict_dfs_INDEX[i] = pd.DataFrame()
            self.dict_dfs_STOCK[i] = pd.DataFrame()
        self.stop: bool = False
        self.curr_time = ""
        self.interval = 5#seconds
        self.red: str = "#e53935"
        self.green: str = "#00e676"
        self.df_INDEX: pd.DataFrame = pd.DataFrame()
        self.df_STOCK: pd.DataFrame = pd.DataFrame()
        self.hdr: Dict[str, str] = {'user-agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 \
                                (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36','accept-language':'en-US,en;q=0.9'}
        self.stock_symbs: List[str] = ['AARTIIND', 'ACC', 'ADANIENT', 'ADANIPORTS', 'AMARAJABAT', 'AMBUJACEM', 'APOLLOHOSP',
                                  'APOLLOTYRE', 'ASHOKLEY',
                                  'ASIANPAINT', 'AUROPHARMA', 'AXISBANK', 'BAJAJ-AUTO', 'BAJAJFINSV', 'BAJFINANCE',
                                  'BALKRISIND', 'BANDHANBNK',
                                  'BANKBARODA', 'BATAINDIA', 'BEL', 'BERGEPAINT', 'BHARATFORG', 'BHARTIARTL', 'BHEL',
                                  'BIOCON', 'BOSCHLTD',
                                  'BPCL', 'BRITANNIA', 'CADILAHC', 'CANBK', 'CHOLAFIN', 'CIPLA', 'COALINDIA', 'COFORGE',
                                  'COLPAL', 'CONCOR',
                                  'CUMMINSIND', 'DABUR', 'DIVISLAB', 'DLF', 'DRREDDY', 'EICHERMOT', 'ESCORTS',
                                  'EXIDEIND', 'FEDERALBNK', 'GAIL',
                                  'GLENMARK', 'GMRINFRA', 'GODREJCP', 'GODREJPROP', 'GRASIM', 'HAVELLS', 'HCLTECH',
                                  'HDFC', 'HDFCAMC', 'HDFCBANK',
                                  'HDFCLIFE', 'HEROMOTOCO', 'HINDALCO', 'HINDPETRO', 'HINDUNILVR', 'IBULHSGFIN',
                                  'ICICIBANK', 'ICICIGI',
                                  'ICICIPRULI', 'IDEA', 'IDFCFIRSTB', 'IGL', 'INDIGO', 'INDUSINDBK', 'INDUSTOWER',
                                  'INFY', 'IOC',
                                  'ITC', 'JINDALSTEL', 'JSWSTEEL', 'JUBLFOOD', 'KOTAKBANK','LALPATHLAB',
                                  'LICHSGFIN', 'LT', 'LUPIN',
                                  'MANAPPURAM', 'MARICO', 'MARUTI', 'MCDOWELL-N', 'MFSL', 'MGL',
                                  'MINDTREE', 'MOTHERSUMI', 'MRF',
                                  'MUTHOOTFIN', 'NATIONALUM', 'NAUKRI', 'NESTLEIND', 'NMDC', 'NTPC', 'ONGC', 'PAGEIND',
                                  'PEL', 'PETRONET', 'PFC',
                                  'PIDILITIND', 'PNB', 'POWERGRID', 'PVR', 'RAMCOCEM', 'RBLBANK', 'RECLTD', 'RELIANCE',
                                  'SAIL', 'SBILIFE', 'SBIN',
                                  'SHREECEM', 'SIEMENS', 'SRF', 'SRTRANSFIN', 'SUNPHARMA', 'SUNTV', 'TATACHEM',
                                  'TATACONSUM', 'TATAMOTORS',
                                  'TATAPOWER', 'TATASTEEL', 'TCS', 'TECHM', 'TITAN', 'TORNTPHARM', 'TORNTPOWER',
                                  'TVSMOTOR', 'UBL', 'ULTRACEMCO',
                                  'UPL', 'VEDL', 'VOLTAS', 'WIPRO']
        #self.stock_symbs: List[str] = ['INFY','UPL', 'VEDL', 'VOLTAS', 'WIPRO', 'ZEEL']
        #self.stock_symbs: List[str] = ['INFY','UPL']
        self.indices: List[str] = ['NIFTY','BANKNIFTY']
        self.SYMBS: List[String] = self.indices
        self.stock_symb: String = ""
        self.session: requests.Session = requests.Session()
        self.make_intial_nse_connection()
        self.expiry_date: String = ""
        self.sheet_window(window)
        self.resposne: requests.Response = requests.Response()
    
    def make_intial_nse_connection(self)->None:
        #if(self.stock_symb==""):
        #    self.session.close()
        self.session.close()
        self.session = requests.Session()
        url_oc = 'https://www.nseindia.com/option-chain'
        self.response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        self.cookies = self.response.cookies
        if(self.stock_symb!=""):
            try:
                self.response = self.session.get(url,headers=self.hdr,timeout=5,cookies=self.cookies)
            except:
                self.make_intial_nse_connection()
    

    def refresh_data(self)->None:
        self.col_time = datetime.now().strftime("%H:%M:%S")
        self.prev_time = time.time()
        
        if(self.sh_frame):
            self.sh_frame.destroy()
        self.sh_frame: Frame = Frame(self.sh_window,height=2)
        self.sh_frame.rowconfigure(0, weight=1)
        self.sh_frame.columnconfigure(0, weight=1)
        self.sh_frame.pack(anchor=N,fill="x", expand=False)
        
        if(self.stock_frame):
            self.stock_frame.destroy()
        self.stock_frame: Frame = Frame(self.sh_window)
        self.stock_frame.rowconfigure(3, weight=1)
        self.stock_frame.columnconfigure(0, weight=1)
        self.stock_frame.pack(anchor=N,fill="both", expand=True)
        
        self.SYMBS = self.indices
        self.index_call = True
        self.set_sheet(self.sh_frame)
        self.sheet_formatting()
        
        self.SYMBS = self.stock_symbs
        self.index_call = False
        self.set_sheet(self.stock_frame)
        self.sheet_formatting()
    
    def set_sheet(self,my_frame: Frame)->None:
        self.NB: Notebook = Notebook(my_frame)
        self.NB.pack(anchor=N,fill="both", expand=True)
        self.NBF: List[Frame] = []
        self.NBS: List[tksheet.Sheet] = []
        self.NB_DF: List[pd.Dataframe] = []
        
        df, dict_dfs = self.append_df_with_OC()
        for key in dict_dfs.keys():
            self.NBF.append(Frame(self.NB))
            self.NB_DF.append(pd.concat([df,dict_dfs[key]],axis = 1))
            self.NB.add(self.NBF[-1],text=key)
            sh = tksheet.Sheet(self.NBF[-1], column_width=70, align="center",
                                                  headers = list(self.NB_DF[-1].columns),
                                                  header_font=("TkDefaultFont", 10, "bold"),
                                                  empty_horizontal=0, empty_vertical=20, header_height=35)
            sh.enable_bindings(
            ("toggle_select", "drag_select", "column_select", "row_select", "column_width_resize",
             "arrowkeys", "right_click_popup_menu", "rc_select", "copy", "select_all"))
            sh.pack(anchor=W,fill="both", expand=True)
            self.NBS.append(sh)
        
    def set_expiry_date(self,event)->None:
        self.expiry_date = self.date_combo_box.get()
    
    def set_ref_intvl(self,event)->None:
        self.interval = float(self.ref_intvl_cbox.get())*60.
    
    def sheet_formatting(self)->None:
        num_std_cols = 2#Symb & ATM
        for i in range(len(self.NBS)):
            curr_sh = self.NBS[i]
            num_cols = len(self.NB_DF[i].columns)
            for col in enumerate(self.NB_DF[i].columns):
                curr_sh.set_column_data(col[0],values=self.NB_DF[i][col[1]])
            if(not self.first_run):
                for i in range(curr_sh.get_total_rows()):
                    for j in range(num_cols-1,num_std_cols,-1):
                        diff = float(curr_sh.get_cell_data(i,j)) - float(curr_sh.get_cell_data(i,j-1))
                        perc_change = 1.
                        if(float(curr_sh.get_cell_data(i,j-1)!=0.)):
                            perc_change = diff*100/float(curr_sh.get_cell_data(i,j-1))
                        if (diff<0.):
                            curr_sh.highlight_cells(row=i, column=j, bg=self.red,fg='white')
                        elif diff==0.0:
                            curr_sh.highlight_cells(row=i, column=j, bg='white',fg='black')
                        else:
                            curr_sh.highlight_cells(row=i, column=j, bg='blue',fg='white')
                        if perc_change>40.:
                            curr_sh.highlight_cells(row=i, column=j, bg='green',fg='white')
            curr_sh.set_currently_selected(0,num_cols-1)
            curr_sh.refresh()
    def sheet_window(self,window)->None:
        self.sh_window: Tk = window
        self.sh_window.title('Option chain analyzer')
        window_width: int = self.sh_window.winfo_reqwidth()
        window_height: int = self.sh_window.winfo_reqheight()
        position_right: int = int(self.sh_window.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.sh_window.winfo_screenheight() / 2 - window_height / 2)
        #self.sh_window.geometry("1200x600+{}+{}".format(position_right, position_down))
        self.sh_window.geometry("1200x600+300+200")
        #self.sh_window.geometry("+{}+{}".format(position_right, position_down))

        self.sh_frame: Frame = Frame(self.sh_window)
        self.stock_frame: Frame = Frame(self.sh_window)
        
        top_frame: Frame = Frame(self.sh_window)
        top_frame.rowconfigure(0, weight=1)
        top_frame.columnconfigure(0, weight=1)
        top_frame.pack(anchor=N, expand=False, side=LEFT)
        
        row_idx: int = 0
        
        self.index_call = True
        self.get_expiry_dates()
        date_var: StringVar = StringVar()
        date_var.set(" ")
        lbl_exp_date: Label = Label(top_frame,text='Index Expiry',justify=LEFT,font=("TkDefaultFont", 10, "bold"))
        lbl_exp_date.grid(row=row_idx,column=0,sticky=N+S+W)
        self.date_combo_box = Combobox(top_frame,width=30,textvariable=date_var) 
        self.date_combo_box.grid(row=row_idx, column=1, sticky=N + S + E + W)
        self.date_combo_box.bind('<<ComboboxSelected>>', self.set_expiry_date)
        self.date_combo_box['values'] = tuple(self.expiry_dates)
        self.date_combo_box.set(self.expiry_dates[0])
        row_idx += 1
        
        self.index_call = False
        self.get_expiry_dates()
        date_var_stock: StringVar = StringVar()
        date_var_stock.set(" ")
        lbl_exp_date_stock: Label = Label(top_frame,text='Stock Expiry',justify=LEFT,font=("TkDefaultFont", 10, "bold"))
        lbl_exp_date_stock.grid(row=row_idx,column=0,sticky=N+S+W)
        self.date_combo_box_stock = Combobox(top_frame,width=30,textvariable=date_var_stock) 
        self.date_combo_box_stock.grid(row=row_idx, column=1, sticky=N + S + E + W)
        self.date_combo_box_stock.bind('<<ComboboxSelected>>', self.set_expiry_date)
        self.date_combo_box_stock['values'] = tuple(self.expiry_dates)
        self.date_combo_box_stock.set(self.expiry_dates[0])
        row_idx += 1
        
        ref_intvl: Stringvar = StringVar()
        ref_intvl.set(" ")
        lbl_refresh_interval: Label = Label(top_frame,text='Interval (min)',justify=LEFT,font=("TkDefaultFont", 10, "bold"))
        lbl_refresh_interval.grid(row=row_idx,column=0,sticky=N+S+W)
        self.ref_intvl_cbox: Combobox = Combobox(top_frame,width=30,textvariable=ref_intvl)
        self.ref_intvl_cbox.grid(row=row_idx, column=1, sticky=N + S + E + W)
        self.ref_intvl_cbox.bind('<<ComboboxSelected>>', self.set_ref_intvl)
        self.ref_intvl_cbox['values'] = tuple(range(1,10,1))
        self.ref_intvl_cbox.configure(state='readonly')
        row_idx += 1

        self.start_button: Button = Button(top_frame,text='START',command=self.main_recursive,width=3)
        self.start_button.grid(row=row_idx, column=1, sticky=N + S + E + W)
        row_idx += 1

        self.sh_window.mainloop()
    
    def get_stock_symbols(self)->None:
        url = 'https://www.nseindia.com/api/master-quote'
        url_oc = 'https://www.nseindia.com/'
        response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        ck = response.cookies
        response = self.session.get(url,headers=self.hdr,timeout=5,cookies=ck)
        json_data = response.json()
        self.stock_symbs = list(json_data)
        
    def get_option_chain_data(self)->None:
        if self.index_call:
            self.url = 'https://www.nseindia.com/api/option-chain-indices?symbol='+self.stock_symb
        else:
            self.url = 'https://www.nseindia.com/api/option-chain-equities?symbol='+self.stock_symb
        try:
            self.response = self.session.get(self.url,headers=self.hdr,timeout=5,cookies=self.cookies)
        except Exception as err:
            print(self.response)
            print(err, "5")
            self.make_intial_nse_connection()
    
    def get_expiry_dates(self)->Tuple:
        if self.index_call:
            self.stock_symb = 'NIFTY'
        else:
            self.stock_symb = 'ASHOKLEY'
        self.get_option_chain_data() 
        json_data = self.response.json()
        self.expiry_dates: List = []
        self.expiry_dates = json_data['records']['expiryDates']
    

    def append_df_with_OC(self):
        self.atms: List = []
        self.pcr: List = []
        self.live_prices: List[float] = []
        self.DIFF_otm: List[float] = []
        self.call_sum_otm: List[float] = []
        self.put_sum_otm: List[float] = []
        self.call_sum_far_otm: List[float] = []
        self.put_sum_far_otm: List[float] = []
        for stk in self.SYMBS:
            self.stock_symb = stk
            self.get_option_chain_data() 
            json_data = self.response.json()
            #print(stk)
            #quote = nse.get_quote(stk)
            #print(quote['data'][0]['lastPrice'])
            #print(json_data['records']['data'])
            #try:
            #    self.live_prices.append(float(quote['data'][0]['lastPrice'].replace(',','')))
            #except:
            #    self.live_prices.append('null')
            if self.index_call:
                match_date = self.date_combo_box.get()
            else:
                match_date = self.date_combo_box_stock.get()
            my_atm: float = 0.0
            strike_prices: List[float] = [data['strikePrice'] for data in json_data['records']['data'] \
                                       if (str(data['expiryDate']).lower() == str(match_date).lower())]
            ce_values: List[dict] = [data['CE'] for data in json_data['records']['data'] \
                        if "CE" in data and (str(data['expiryDate'].lower()) == str(match_date.lower()))]
            pe_values: List[dict] = [data['PE'] for data in json_data['records']['data'] \
                        if "PE" in data and (str(data['expiryDate'].lower()) == str(match_date.lower()))]
            
            ce_data: pd.DataFrame = pd.DataFrame(ce_values)
            pe_data: pd.DataFrame = pd.DataFrame(pe_values)


            curr_price = ce_data['underlyingValue'][0]
            self.live_prices.append(curr_price)
            try:
                diff = [abs(x-curr_price) for x in strike_prices]
                min_pos = diff.index(min(diff))
                my_atm = strike_prices[min_pos]
                self.atms.append(my_atm)
            except:
                self.atms.append('null')
            
            ce_data_otm: pd.DataFrame = ce_data[min_pos+1:min_pos+4]
            pe_data_otm: pd.DataFrame = pe_data[min_pos-3:min_pos]
            
            ce_data_far_otm: pd.DataFrame = ce_data[min_pos+4:min_pos+8]
            pe_data_far_otm: pd.DataFrame = pe_data[min_pos-7:min_pos-3]

            call_sum_otm =  ce_data_otm['changeinOpenInterest'].sum()
            put_sum_otm  =  pe_data_otm['changeinOpenInterest'].sum()

            diff_otm = call_sum_otm - put_sum_otm 
            
            call_sum_far_otm =  ce_data_far_otm['changeinOpenInterest'].sum()
            put_sum_far_otm =  pe_data_far_otm['changeinOpenInterest'].sum()

            diff_far_otm = call_sum_far_otm - put_sum_far_otm

            if(ce_data['openInterest'].sum()!=0): 
                pcr_ratio = float(pe_data['openInterest'].sum())/float(ce_data['openInterest'].sum())
            else:
                pcr_ratio = float(pe_data['openInterest'].sum() )/.001
            self.pcr.append(pcr_ratio)
            self.DIFF_otm.append(diff_otm)
            self.call_sum_otm.append(call_sum_otm)
            self.put_sum_otm.append(put_sum_otm)
            self.call_sum_far_otm.append(call_sum_far_otm)
            self.put_sum_far_otm.append(put_sum_far_otm)
        
        if(self.index_call):
            self.df_INDEX['SYMB'] = self.SYMBS
            self.df_INDEX['ATM'] = self.atms
            self.dict_dfs_INDEX['PCR'][self.col_time] = self.pcr 
            self.dict_dfs_INDEX['PCR'][self.col_time] = self.dict_dfs_INDEX['PCR'][self.col_time].round(3)
            self.dict_dfs_INDEX['CE-OTM'][self.col_time] = self.call_sum_otm
            self.dict_dfs_INDEX['PE-OTM'][self.col_time] = self.put_sum_otm
            self.dict_dfs_INDEX['CE-FAR_OTM'][self.col_time] = self.call_sum_far_otm
            self.dict_dfs_INDEX['PE-FAR_OTM'][self.col_time] = self.put_sum_far_otm
            self.dict_dfs_INDEX['LTP'][self.col_time] = self.live_prices
            return self.df_INDEX, self.dict_dfs_INDEX
        else:
            self.df_STOCK['SYMB'] = self.SYMBS
            self.df_STOCK['ATM'] = self.atms
            self.dict_dfs_STOCK['PCR'][self.col_time] = self.pcr 
            self.dict_dfs_STOCK['PCR'][self.col_time] = self.dict_dfs_STOCK['PCR'][self.col_time].round(3)
            self.dict_dfs_STOCK['CE-OTM'][self.col_time] = self.call_sum_otm
            self.dict_dfs_STOCK['PE-OTM'][self.col_time] = self.put_sum_otm
            self.dict_dfs_STOCK['CE-FAR_OTM'][self.col_time] = self.call_sum_far_otm
            self.dict_dfs_STOCK['PE-FAR_OTM'][self.col_time] = self.put_sum_far_otm
            self.dict_dfs_STOCK['LTP'][self.col_time] = self.live_prices
            return self.df_STOCK, self.dict_dfs_STOCK
                    
if __name__ == '__main__':
    master_window: Tk = Tk()
    NSE(master_window)
    master_window.mainloop()


