#!/usr/bin/python3

import requests
from typing import Union, Optional, List, Dict, Tuple, TextIO, Any
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, \
    Label, Entry, SOLID, RIDGE, N, S, E, W, LEFT, messagebox
from tkinter.ttk import Combobox, Button
import pandas as pd

class NSE:
    def __init__(self,window: Tk)->None:
        self.hdr: Dict[str, str] = {'user-agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 \
                                (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36','accept-language':'en-US,en;q=0.9'}
        self.session = requests.Session()
        self.get_stock_symbols()
        self.login_window(window)
        self.expiry_date: String = ""
        self.stock_symb: String = ""
        self.resposne: requests.Response = requests.Response()
    
    def set_expiry_date(self,event)->None:
        self.expiry_date = self.date_combo_box.get()
    
    def set_stock_symb(self,event)->None:
        self.stock_symb = self.stock_combo_box.get() 
    
    def dummy(self,event)->None:
        print("Dummy\n")
    
    def login_window(self,window)->None:
        self.lw: Tk = window
        self.lw.title('Option chain analyzer')
        window_width: int = self.lw.winfo_reqwidth()
        window_height: int = self.lw.winfo_reqheight()
        position_right: int = int(self.lw.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.lw.winfo_screenheight() / 2 - window_height / 2)
        self.lw.geometry("400x150+{}+{}".format(position_right, position_down))
        
        self.lw.rowconfigure(0, weight=1)
        self.lw.rowconfigure(1, weight=1)
        self.lw.rowconfigure(2, weight=1)
        self.lw.rowconfigure(3, weight=1)
        self.lw.columnconfigure(0, weight=1)
        self.lw.columnconfigure(1, weight=1)
        self.lw.columnconfigure(2, weight=1)
        
        stock_symb_var: StringVar = StringVar()
        stock_symb_var.set("")
        
        date_var: StringVar = StringVar()
        date_var.set("")
        symb_var: StringVar = StringVar()
        symb_var.set("")
        self.data_get: Button = Button(self.lw,text='Fetch data',command=self.get_dataframe,width=1)
        self.data_get.grid(row=3, column=2, sticky=N + S + E + W)
        
        lbl_stock_symb: Label = Label(self.lw,text='Stock symbol',justify=LEFT)
        lbl_stock_symb.grid(row=1,column=0,sticky=N+S+W)
        self.stock_combo_box = Combobox(self.lw,width=5,textvariable=stock_symb_var)
        self.stock_combo_box.grid(row=1, column=1, sticky=N + S + E + W)
        self.stock_combo_box['values'] = self.stock_symbs
        self.stock_combo_box.bind('<<ComboboxSelected>>', self.get_data)
        
        lbl_exp_date: Label = Label(self.lw,text='Expiry date',justify=LEFT)
        lbl_exp_date.grid(row=2,column=0,sticky=N+S+W)
        self.date_combo_box = Combobox(self.lw,width=3,textvariable=date_var) 
        self.date_combo_box.grid(row=2, column=1, sticky=N + S + E + W)
        self.date_combo_box.bind('<<ComboboxSelected>>', self.set_expiry_date)

        lbl_strike_prices: Lable = Label(self.lw,text='Strike prices',justify=LEFT)
        lbl_strike_prices.grid(row=3,column=0,sticky=N+S+W)
        self.strike_combo_box = Combobox(self.lw,width=3,textvariable=date_var) 
        self.strike_combo_box.grid(row=3, column=1, sticky=N + S + E + W)
        #self.strike_combo_box.bind('<<ComboboxSelected>>', self.set_expiry_date)
        self.strike_combo_box['values'] = self.strike_prices
        
        self.lw.mainloop()
    
    def get_stock_symbols(self)->None:
        url = 'https://www.nseindia.com/api/master-quote'
        url_oc = 'https://www.nseindia.com/'
        response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        ck = response.cookies
        response = self.session.get(url,headers=self.hdr,timeout=5,cookies=ck)
        json_data = response.json()
        self.stock_symbs = tuple(json_data)
        
    #def get_data(self,event)->requests.Response:
    def get_data(self,event)->None:
        self.stock_symb = self.stock_combo_box.get()
        url = 'https://www.nseindia.com/api/option-chain-equities?symbol='+self.stock_symb
        url_oc = 'https://www.nseindia.com/option-chain'
        response = self.session.get(url_oc,headers=self.hdr,timeout=5)
        ck = response.cookies
        response = self.session.get(url,headers=self.hdr,timeout=5,cookies=ck)
        self.response = response
        self.get_expiry_dates()
        #json_data = response.json()
        #print(json_data)
    
    def get_expiry_dates(self)->Tuple:
        json_data = self.response.json()
        expiry_dates: List = []
        for i in range(len(json_data['records']['data'])):
            dd = json_data['records']['data'][i]['expiryDate']
            if dd not in expiry_dates:
                expiry_dates.append(json_data['records']['data'][i]['expiryDate'])
        self.date_combo_box['values'] = tuple(expiry_dates)
        self.date_combo_box.set(expiry_dates[0])
    
    def get_dataframe(self)->None:
        #response = self.get_data()
        json_data = response.json()
        df: pd.DataFrame = pd.read_json(response.text)
        
        strikes = []
        #print((json_data['records']['data']).keys())
        ce_data: List = [data['CE'] for data in json_data['records']['data'] if "CE" in data]
        pe_data: List = [data['PE'] for data in json_data['records']['data'] if "PE" in data]

        ce_values: pd.DataFrame = pd.DataFrame(ce_data)
        pe_values: pd.DataFrame = pd.DataFrame(pe_data)

        print(ce_data)
                    


if __name__ == '__main__':
    master_window: Tk = Tk()
    NSE(master_window)
    master_window.mainloop()


