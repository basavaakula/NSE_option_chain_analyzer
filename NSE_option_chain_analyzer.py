#!/usr/bin/python3

import requests
from typing import Union, Optional, List, Dict, Tuple, TextIO, Any
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, \
    Label, Entry, SOLID, RIDGE, N, S, E, W, LEFT, messagebox
from tkinter.ttk import Combobox, Button
import pandas as pd



class NSE:
    
    def __init__(self,window: Tk)->None:
        self.login_window(window)
    
    def login_window(self,window)->None:
        self.lw: Tk = window
        self.lw.title('Option chain analyzer')
        window_width: int = self.lw.winfo_reqwidth()
        window_height: int = self.lw.winfo_reqheight()
        position_right: int = int(self.lw.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.lw.winfo_screenheight() / 2 - window_height / 2)
        self.lw.geometry("320x110+{}+{}".format(position_right, position_down))
        
        self.lw.rowconfigure(0, weight=1)
        self.lw.rowconfigure(1, weight=1)
        self.lw.rowconfigure(2, weight=1)
        self.lw.rowconfigure(3, weight=1)
        self.lw.columnconfigure(0, weight=1)
        self.lw.columnconfigure(1, weight=1)
        self.lw.columnconfigure(2, weight=1)

        self.get_data: Button = Button(self.lw,text='Fetch data',command=self.get_dataframe,width=100)
        self.get_data.grid(row=0, column=2, sticky=N + S + E + W)
        self.lw.mainloop()

    #def get_data(self)->Optional[requests.Response,Any]:
    def get_data(self)->requests.Response:
        session = requests.Session()

        url = 'https://www.nseindia.com/api/option-chain-equities?symbol=ASHOKLEY'
        url_oc = 'https://www.nseindia.com/option-chain'

        hdr: Dict[str, str] = {'user-agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36','accept-language':'en-US,en;q=0.9'}

#hdr = {'user-agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36'}
#hdr ={'accept-language':'en-US,en;q=0.9'}
#tt = ['Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/53    7.36 (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36','accept-language':'en-US,en;q=0.9']

        response = session.get(url_oc,headers=hdr,timeout=5)
        ck = response.cookies
        response = session.get(url,headers=hdr,timeout=5,cookies=ck)

        json_data = response.json()
        return response
    def get_dataframe(self)->None:
        #self.response,self.json_data = self.get_data()
        response = self.get_data()
        df: pd.DataFrame = pd.read_json(response.text)
        print(df)


if __name__ == '__main__':
    master_window: Tk = Tk()
    NSE(master_window)
    master_window.mainloop()


