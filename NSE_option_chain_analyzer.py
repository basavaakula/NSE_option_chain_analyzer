#!/usr/bin/python3

import requests
from typing import Union, Optional, List, Dict, Tuple, TextIO, Any
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, \
    Label, Entry, SOLID, RIDGE, N, S, E, W, LEFT, messagebox
from tkinter.ttk import Combobox, Button
import pandas as pd
import tksheet
import nsepy.live as nse
from datetime import datetime

class NSE:
    def __init__(self,window: Tk)->None:
        self.red: str = "#e53935"
        self.green: str = "#00e676"
        self.df: pd.DataFrame = pd.DataFrame()
        self.sheet_col_hdrs: Tuple[str,str] = ('Stocks','ATM')
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
        #self.stock_symbs: List[str] = ['ZEEL']
        self.session = requests.Session()
        self.df['Stocks'] = self.stock_symbs
        self.expiry_date: String = ""
        self.sheet_window(window)
        self.stock_symb: String = ""
        self.resposne: requests.Response = requests.Response()
    
    def refresh_data(self)->None:
        timee = datetime.now().strftime("%H:%M:%S") 
        #self.get_live_stock_price()
        self.append_df_with_OC()
        self.populate_sheet1()
        self.stock_combo_box.configure(state='readonly')
        for i in range(self.sheet.get_total_rows()):
            pcr = float(self.sheet.get_cell_data(i,3))
            if(pcr>1.):
                self.sheet.highlight_cells(row=i, column=3, bg=self.green)
            else:
                self.sheet.highlight_cells(row=i, column=3, bg=self.red)
        self.sh_window.mainloop()
    
    def set_sheet(self)->None:
        #self.df = self.df[self.df['Current Price'].notna()]
        if(self.sh_frame):
            self.sh_frame.destroy()
        self.sh_frame: Frame = Frame(self.sh_window)
        self.sh_frame.rowconfigure(2, weight=1)
        self.sh_frame.columnconfigure(0, weight=1)
        self.sh_frame.pack(anchor=N,fill="both", expand=True)
        self.sheet: tksheet.Sheet = tksheet.Sheet(self.sh_frame, column_width=110, align="center",
                                                  headers = list(self.df.columns), 
                                                  header_font=("TkDefaultFont", 10, "bold"),
                                                  empty_horizontal=0, empty_vertical=20, header_height=35)
        self.sheet.enable_bindings(
            ("toggle_select", "drag_select", "column_select", "row_select", "column_width_resize",
             "arrowkeys", "right_click_popup_menu", "rc_select", "copy", "select_all"))
        self.sheet.pack(anchor=W,fill="both", expand=True)
    
    def set_expiry_date(self,event)->None:
        self.expiry_date = self.date_combo_box.get()
    
    def set_stock_symb(self,event)->None:
        self.stock_symb = self.stock_combo_box.get() 
    
    def dummy(self,event)->None:
        print("Dummy\n")
    
    def populate_sheet(self,event)->None:
        sub_df: pd.DataFrame = pd.DataFrame()   
        stock_selected: String = self.stock_combo_box.get()
        if(stock_selected!=" "):
            sub_df = self.df.loc[self.df['Stocks'] == stock_selected]
        else:
            sub_df = self.df
        
        self.set_sheet()
        for col in enumerate(sub_df.columns):
            self.sheet.set_column_data(col[0],values=sub_df[col[1]])

    
    def populate_sheet1(self)->None:
        sub_df: pd.DataFrame = pd.DataFrame()   
        stock_selected: String = self.stock_combo_box.get()
        if(stock_selected!=" "):
            sub_df = self.df.loc[self.df['Stocks'] == stock_selected]
        else:
            sub_df = self.df
        
        self.set_sheet()
        for col in enumerate(sub_df.columns):
            self.sheet.set_column_data(col[0],values=sub_df[col[1]])
    
    def get_live_stock_price(self)->None:
        self.live_prices: List[float] = []
        for stk in self.stock_symbs:
            quote = nse.get_quote(stk)
            try:
                self.live_prices.append(float(quote['data'][0]['lastPrice'].replace(',','')))
            except:
                self.live_prices.append('null')
        self.df['Current Price'] = self.live_prices
    
    def sheet_window(self,window)->None:
        self.get_expiry_dates()
        self.sh_window: Tk = window
        self.sh_window.title('Option chain analyzer')
        window_width: int = self.sh_window.winfo_reqwidth()
        window_height: int = self.sh_window.winfo_reqheight()
        position_right: int = int(self.sh_window.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.sh_window.winfo_screenheight() / 2 - window_height / 2)
        self.sh_window.geometry("1600x800+{}+{}".format(position_right, position_down))

        self.sh_frame: Frame = Frame(self.sh_window)
        
        top_frame: Frame = Frame(self.sh_window)
        top_frame.rowconfigure(0, weight=1)
        top_frame.columnconfigure(0, weight=1)
        top_frame.pack(anchor=N, expand=False, side=LEFT)
        #top_frame.grid(row=0, column=0, sticky=N + S + W + E)
        
        stock_symb_var: StringVar = StringVar()
        stock_symb_var.set(" ")
        lbl_stock_symb: Label = Label(top_frame,text='Stock symbol',justify=LEFT,font=("TkDefaultFont", 10, "bold"))
        lbl_stock_symb.grid(row=0,column=0,sticky=N+S+W)
        self.stock_combo_box = Combobox(top_frame,width=30,textvariable=stock_symb_var)
        self.stock_combo_box.grid(row=0, column=1, sticky=N + S + E + W)
        self.stock_combo_box['values'] = self.stock_symbs
        self.stock_combo_box.bind('<<ComboboxSelected>>', self.populate_sheet)
        self.stock_combo_box.configure(state='disabled')
        
        date_var: StringVar = StringVar()
        date_var.set(" ")
        lbl_exp_date: Label = Label(top_frame,text='Expiry date',justify=LEFT,font=("TkDefaultFont", 10, "bold"))
        lbl_exp_date.grid(row=2,column=0,sticky=N+S+W)
        self.date_combo_box = Combobox(top_frame,width=30,textvariable=date_var) 
        self.date_combo_box.grid(row=2, column=1, sticky=N + S + E + W)
        self.date_combo_box.bind('<<ComboboxSelected>>', self.set_expiry_date)
        self.date_combo_box['values'] = tuple(self.expiry_dates)
        self.date_combo_box.set(self.expiry_dates[0])
        
        self.start_button: Button = Button(top_frame,text='START',command=self.refresh_data,width=3)
        self.start_button.grid(row=4, column=1, sticky=N + S + E + W)
        
        self.sh_window.mainloop()
        
    def get_stock_symbols(self)->None:
        url = 'https://www.nseindia.com/api/master-quote'
        url_oc = 'https://www.nseindia.com/'
        response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        ck = response.cookies
        response = self.session.get(url,headers=self.hdr,timeout=5,cookies=ck)
        json_data = response.json()
        self.stock_symbs = list(json_data)
        
    #def get_option_chain_data(self,event)->requests.Response:
    def get_option_chain_data(self)->None:
        url = 'https://www.nseindia.com/api/option-chain-equities?symbol='+self.stock_symb
        url_oc = 'https://www.nseindia.com/option-chain'
        response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        ck = response.cookies
        response = self.session.get(url,headers=self.hdr,timeout=5,cookies=ck)
        self.response = response
        #self.get_expiry_dates()
    
    def get_expiry_dates(self)->Tuple:
        self.stock_symb = "ASHOKLEY"
        self.get_option_chain_data() 
        json_data = self.response.json()
        self.expiry_dates: List = []
        self.expiry_dates = json_data['records']['expiryDates']
        #for i in range(len(json_data['records']['data'])):
        #    dd = json_data['records']['data'][i]['expiryDate']
        #    if dd not in self.expiry_dates:
        #        self.expiry_dates.append(json_data['records']['data'][i]['expiryDate'])
    
    def append_df_with_OC(self)->None:
        self.atms: List = []
        self.pcr: List = []
        self.live_prices: List[float] = []
        for stk in self.stock_symbs:
            self.stock_symb = stk
            self.get_option_chain_data() 
            json_data = self.response.json()
            print(stk)
            #print(json_data)
            quote = nse.get_quote(stk)
            try:
                self.live_prices.append(float(quote['data'][0]['lastPrice'].replace(',','')))
            except:
                self.live_prices.append('null')
            my_atm: float = 0.0
            strike_prices: List[float] = [data['strikePrice'] for data in json_data['records']['data'] \
                                       if (str(data['expiryDate']).lower() == str(self.date_combo_box.get()).lower())]
            try:
                curr_price = float(quote['data'][0]['lastPrice'].replace(',',''))
                diff = [abs(x-curr_price) for x in strike_prices]
                min_pos = diff.index(min(diff))
                my_atm = strike_prices[min_pos]
                self.atms.append(my_atm)
            except:
                self.atms.append('null')
            ce_values: List[dict] = [data['CE'] for data in json_data['records']['data'] \
                        if "CE" in data and (str(data['expiryDate'].lower()) == str(self.date_combo_box.get().lower()))]
            pe_values: List[dict] = [data['PE'] for data in json_data['records']['data'] \
                        if "PE" in data and (str(data['expiryDate'].lower()) == str(self.date_combo_box.get().lower()))]
            
            ce_data: pd.DataFrame = pd.DataFrame(ce_values)
            pe_data: pd.DataFrame = pd.DataFrame(pe_values)
            
            ce_data_sub: pd.DataFrame = ce_data[min_pos+1:min_pos+4]
            pe_data_sub: pd.DataFrame = pe_data[min_pos-3:min_pos]

            #print(ce_data_sub)
            #print(pe_data_sub)
            
            #print(my_atm)
            #print(pe_data_sub['openInterest'].sum())
            #print(ce_data_sub['openInterest'].sum())
            if(ce_data_sub['openInterest'].sum()!=0): 
                pcr_ratio = float(pe_data_sub['openInterest'].sum())/float(ce_data_sub['openInterest'].sum())
            else:
                pcr_ratio = float(pe_data_sub['openInterest'].sum())/.01
            self.pcr.append(pcr_ratio)
        
        self.df['Current Price'] = self.live_prices
        self.df['ATM'] = self.atms
        self.df['PCR'] = self.pcr
        self.df['PCR'] = self.df['PCR'].round(3)
            #try:
            #    self.live_prices.append(quote['data'][0]['lastPrice'])
            #except:
            #    self.live_prices.append('null')
        #self.df['Current Price'] = self.live_prices
        #strikes = []
        ##print((json_data['records']['data']).keys())
        #ce_data: List = [data['CE'] for data in json_data['records']['data'] if "CE" in data]
        #pe_data: List = [data['PE'] for data in json_data['records']['data'] if "PE" in data]

        #ce_values: pd.DataFrame = pd.DataFrame(ce_data)
        #pe_values: pd.DataFrame = pd.DataFrame(pe_data)

        #print(ce_data)
                    
if __name__ == '__main__':
    master_window: Tk = Tk()
    NSE(master_window)
    master_window.mainloop()


