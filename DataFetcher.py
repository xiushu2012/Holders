# -*- coding: utf-8 -*-
import datetime
import random
import sys
import time
import openpyxl
import fire
import pandas as pd
import requests
import os


class TopTheHoldingV2:

    def __init__(self,selected):
        self.today = datetime.datetime.today().strftime('%Y-%m-%d')
        selected_df = pd.read_excel(selected)
        print(selected_df)
        self.jsl_data = [(str(code)[2:],name) for code,name in zip(selected_df['code'], selected_df['name'])]

    @property
    def headers(self):
        return {
            'origin': 'https://emh5.eastmoney.com',
            'pragma': 'no-cache',
            'referer': 'https://emh5.eastmoney.com',
            'user-agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1'
        }

    def crawl(self, code):
        if code.startswith('12'):
            code = code + '.SZ'
        else:
            code = code + '.SH'

        url = 'https://datacenter.eastmoney.com/securities/api/data/get?client=APP&source=SECURITIES&type=RPT_BOND10_BS_HOLDER&sty=SECUCODE,SECURITY_CODE,BOND_NAME_ABBR,HOLDER_NAME,END_DATE,HOLD_NUM,HOLD_RATIO,HOLDER_RANK&filter=(SECUCODE%3D%22{}%22)&pageNumber=1&pageSize=50'.format(
            code)
        print(url)
        response = requests.get(url, headers=self.headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Failed to fetch data for {code}, status code: {response.status_code}")
            return None

    def parse_json(self, js):
        result_list = []

        if js is None or js.get('result') is None:
            return result_list

        result_list = js.get('result', {}).get('data', [])

        return result_list

    def run(self):
        print('in run...')
        for code,name in self.jsl_data:
            print(f'crawling {code}:{name}')
            js = self.crawl(code)
            result_list = self.parse_json(js)
            self.dump_excel(result_list, name) 


    def dump_excel(self,result_list, name):
        if len(result_list) == 0:
            print('empty')
            return

        df = pd.DataFrame(result_list)
        df = df[['BOND_NAME_ABBR', 'SECUCODE', 'END_DATE', 'HOLDER_NAME', 'HOLD_NUM', 'HOLD_RATIO', 'HOLDER_RANK']]


        filename = f'fetch-{self.today}.xlsx'
        try:
            if not os.path.exists(filename):
                # 如果文件不存在，直接写入 DataFrame
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='holders', index=False)
            else:
                # 如果文件存在，读取现有的工作表数据
                with pd.ExcelFile(filename, engine='openpyxl') as xls:
                    existing_df = pd.read_excel(xls, sheet_name='holders')
        
                # 将新的数据追加到现有数据
                combined_df = pd.concat([existing_df, df], ignore_index=True)
    
                # 重新写入工作表
                with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
                    combined_df.to_excel(writer, sheet_name='holders', index=False)            
            print(f'Excel file "{filename}" exported successfully with data for bond {name}.')
        except FileNotFoundError:
            print("Error while writing to Excel file.")

def main(code=None):
    from sys import argv
    filein = ""
    if len(argv) > 1:
        filein = argv[1]
    else:
        print("please run like 'python redeem.py [file]'")
        exit(1)
    app = TopTheHoldingV2(filein)
    app.run()  
if __name__ == '__main__':
    fire.Fire(main)
    
    
    
    
    
  
    
    
    
    