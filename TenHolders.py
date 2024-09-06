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
import re
import matplotlib.pyplot as plt  
# 支持中文
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']

class TopTheHoldingV2:

    def __init__(self,selected):
        self.today = datetime.datetime.today().strftime('%Y-%m-%d')
        selected_df = pd.read_excel(selected)
        print(selected_df)
        self.jsl_data = [(str(code)[2:],name) for code,name in zip(selected_df['code'], selected_df['name'])]
        self.mapper_list =['林园','宁泉','甄投','明汯','汇添富','博时',
                     '易方达','全国社保','兴全','东方红','南方东英','嘉实','富国','天弘',
                     '光大保德','诺安','中欧','中邮','上海迎水','广发','鹏华','上海泉汐','上海睿郡']

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
        df.columns = ['转债名称', '转债代码', '公布日期', '持有人', '持有张数', '持有比例', '排名']


        filename = f'{self.today}.xlsx'
        try:
            if not os.path.exists(filename):
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=name, index=False)
            else:
                with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name=name, index=False)
            
            self.figure_holder(df,name,self.today)
            print(f'Excel file "{filename}" exported successfully with data for bond {name}.')
        except FileNotFoundError:
            print("Error while writing to Excel file.")

    def map_short_name(self,x):
        if len(x)<5:
            return x
        for i in self.mapper_list:
            if re.search(i,x):
                return i
        return x

    def figure_holder(self,df,name,folder):
        isExist = os.path.exists(folder)
        if not isExist:
            os.makedirs(folder)
            print("create figure folder:%s create" % (folder))
                                  
        df['公布日期'] = pd.to_datetime(df['公布日期'])
        #df['持有人']=df['持有人'].map(self.map_short_name)
        
        #这块存在持有人过多几乎无法分清....
        values_to_match = ['合计', '银行','证券','基金','UBS','有限公司','信托']
        pattern = '|'.join(values_to_match)
        df_filtered  = df[~(df['持有人'].str.contains(pattern, regex=True, case=False))]
        plt.figure(figsize=(10, 6))                                                 
        
        #获取唯一持有人列表并为每个持有人分配一个编号
        #df_filtered = df_filtered.sort_values(by='公布日期')
        holders = df_filtered['持有人'].unique()
        holder_dict = {holder: idx + 1 for idx, holder in enumerate(holders)}
        # 绘制每个持有人的数据曲线
        for holder, idx in holder_dict.items():
            holder_data = df_filtered[df_filtered['持有人'] == holder]
            plt.plot(holder_data['公布日期'], holder_data['持有比例'], label=f'{idx}: {holder}')
            
            # 在曲线的最后一个数据点上标注编号
            last_date = holder_data['公布日期'].iloc[0]
            last_ratio = holder_data['持有比例'].iloc[0]
            plt.text(last_date, last_ratio, f'{idx}', fontsize=10, ha='left', va='center')
 
        # 设置图例位置
        plt.legend(loc='best')                                                     
        plt.title('持有人持有比例变化')                           
        plt.xlabel('公布日期')                                                      
        plt.ylabel('持有比例')                                                      
                                                                                    
        # 显示图例                                                                  
        plt.legend(title='持有人', bbox_to_anchor=(1.05, 1), loc='upper left')  
        #plt.legend(title='持有人', loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=3)    
                                                                                    
        # 旋转x轴的日期标签，避免重叠                                               
        plt.xticks(rotation=45)                                                     
                                                                                    
        # 调整布局以适应所有内容                                                    
        plt.tight_layout()                                                          
                                                                                    
        # 显示图形
        figurepath = f'{folder}/{name}.png'
        plt.savefig(figurepath)                                                                  
        #plt.show()    


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
    
    
    
    
    
  
    
    
    
    