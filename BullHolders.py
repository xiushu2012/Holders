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
#import matplotlib.dates as mdates
import networkx as nx
import matplotlib.pyplot as plt
import numpy as np
# 支持中文
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']

class BullHolding:

    def __init__(self,selected,timepoint):
        self.today = datetime.datetime.today().strftime('%Y-%m-%d')
        self.filename = selected
        self.current = timepoint
        self.mapper_list =['林园','宁泉','甄投','明汯','汇添富','博时',
                     '易方达','全国社保','兴全','东方红','南方东英','嘉实','富国','天弘',
                     '光大保德','诺安','中欧','中邮','上海迎水','广发','鹏华','上海泉汐','上海睿郡']


    def run(self):
        print('in run...')
        self.group_excel() 

    def category(self,x):
        # category_list = ['私募','']
        if re.search('私募',x):
            return '私募'
        if len(x)<5:
            return '个人'
        if re.search('社保',x):
            return '社保'
        return '公募'
        

    def figure_network(self, bullholders_df):
        G = nx.Graph()
    
        # 添加节点和边
        for index, row in bullholders_df.iterrows():
            bond = row['BOND_NAME_ABBR']
            holder = row['HOLDER_NAME']
    
            # 添加节点
            if bond not in G.nodes():
                G.add_node(bond, shape='s')  # 方形节点
            if holder not in G.nodes():
                G.add_node(holder, shape='o')  # 圆形节点
    
            # 添加边
            G.add_edge(bond, holder)
    
        # 找到度小于2的持有人点并移除
        holder_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 'o']
        holder_nodes_to_remove = [node for node in holder_nodes if G.degree(node) < 2]
        G.remove_nodes_from(holder_nodes_to_remove)
        
        # 找到度等于0的债券节点并移除
        bond_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 's']
        bond_nodes_to_remove = [node for node in bond_nodes if G.degree(node) < 1]
        G.remove_nodes_from(bond_nodes_to_remove)
    
        # 使用 spring 布局
        #pos = nx.spring_layout(G)
        #pos = nx.circular_layout(G)
        #pos = nx.random_layout(G)

        
        #外圈位置布局均匀分布角度
        bond_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 's']
        num_bonds = len(bond_nodes);out_radius = 400
        bond_theta = np.linspace(0, 2*np.pi, num_bonds, endpoint=False)
        print(bond_theta,num_bonds)
        bond_pos = {bond_nodes[i]: (out_radius * np.cos(bond_theta[i]), out_radius * np.sin(bond_theta[i])) for i in range(num_bonds)}
 

        
        #内圈位置布局均匀分布角度
        holder_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 'o']
        num_holders = len(holder_nodes);inner_radius = 200
        theta = np.linspace(0, 2*np.pi, num_holders, endpoint=False)
        print(theta,num_holders)
        holder_pos = {holder_nodes[i]: (inner_radius * np.cos(theta[i]), inner_radius * np.sin(theta[i])) for i in range(num_holders)}
        
        pos = {**bond_pos, **holder_pos}
    
        # 绘制边
        nx.draw_networkx_edges(G, pos, width=2, alpha=0.5)
    
        # 绘制节点标签,创建一个字典，将节点映射到其名称
        #labels = {node: node for node in G.nodes()} 
        #print(labels)
        labels = {node:  node[:2] if '转债' in node else node for node in G.nodes()}
        
        nx.draw_networkx_labels(G, pos, labels=labels, font_size=10, font_color='darkred')
    
        # 分开处理方形节点和圆形节点
        bond_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 's']
        holder_nodes = [node for node in G.nodes() if G.nodes[node].get('shape') == 'o']
    
        # 调整节点大小
        node_size = 700
        # 绘制方形节点
        nx.draw_networkx_nodes(G, pos, nodelist=bond_nodes, node_color='skyblue', node_shape='s', node_size=node_size)
    
        # 绘制剩余的圆形节点
        nx.draw_networkx_nodes(G, pos, nodelist=holder_nodes, node_color='lightgreen', node_shape='o', node_size=node_size)
    
        # 显示图
        #plt.axis('off')
        #figurepath = f'./holders_network.png'
        #plt.savefig(figurepath) 
        plt.show()

        
    def group_holders(self,bullholders_df):
        #按照'HOLDER_NAME'进行分组，并将每个组的'BOND_NAME_ABBR'聚合成一个列表
        grouped_bullholders = bullholders_df.groupby('HOLDER_NAME')['BOND_NAME_ABBR'].apply(list)


        # 将聚合结果转换为DataFrame
        grouped_bullholders_df = grouped_bullholders.reset_index()
        grouped_bullholders_df['BOND_COUNT'] = grouped_bullholders_df['BOND_NAME_ABBR'].apply(len)
        grouped_bullholders_df = grouped_bullholders_df.sort_values(by='BOND_COUNT', ascending=False)
        grouped_bullholders_df.reset_index(drop=True, inplace=True)
                
        print(grouped_bullholders_df)
        output_filename = self.filename.replace('fetch','fetch-hold')
        grouped_bullholders_df.to_excel(output_filename, index=False, engine='openpyxl')
        
    def group_bonds(self,bullholders_df):
        #按照'BOND_NAME_ABBR'进行分组，并将每个组的'HOLDER_NAME'聚合成一个列表
        grouped_bullholders = bullholders_df.groupby('BOND_NAME_ABBR')['HOLDER_NAME'].apply(list)


        # 将聚合结果转换为DataFrame
        grouped_bullholders_df = grouped_bullholders.reset_index()
        grouped_bullholders_df['HOLDER_COUNT'] = grouped_bullholders_df['HOLDER_NAME'].apply(len)
        grouped_bullholders_df = grouped_bullholders_df.sort_values(by='HOLDER_COUNT', ascending=False)
        grouped_bullholders_df.reset_index(drop=True, inplace=True)
                
        print(grouped_bullholders_df)
        output_filename = self.filename.replace('fetch','fetch-bond')
        grouped_bullholders_df.to_excel(output_filename, index=False, engine='openpyxl')
    

    def group_excel(self):
        #df = df[['BOND_NAME_ABBR', 'SECUCODE', 'END_DATE', 'HOLDER_NAME', 'HOLD_NUM', 'HOLD_RATIO', 'HOLDER_RANK']]
        #df.columns = ['转债名称', '转债代码', '公布日期', '持有人', '持有张数', '持有比例', '排名']


        filename = self.filename
        try:
            if not os.path.exists(filename):
                print(f'Excel file "{filename}" not Exist')
            else:
                print(f'Excel file "{filename}" is Exist.')
                whole_df = pd.read_excel(filename)
                # 确保'END_DATE'列是日期格式并筛选日期大于'2024-06-30'的行
                whole_df['END_DATE'] = pd.to_datetime(whole_df['END_DATE'])
                original_df = whole_df[ (whole_df['END_DATE'] >= self.current) & (whole_df['HOLDER_NAME'] != '合计')]
                
                current_df = original_df.copy()
                #current_df['CATEGORY']=current_df['HOLDER_NAME'].map(self.category)
                current_df.loc[:, 'CATEGORY'] = current_df['HOLDER_NAME'].map(self.category)
                bullholders_df =  current_df[current_df['CATEGORY'] == '个人' ]
                
                self.figure_network(bullholders_df)
                self.group_holders(bullholders_df)
                self.group_bonds(bullholders_df)
                
        except FileNotFoundError:
            print("Error while reading Excel file.")

    def map_short_name(self,x):
        if len(x)<5:
            return x
        for i in self.mapper_list:
            if re.search(i,x):
                return i
        return x


def main(code=None):
    from sys import argv
    filein = ""
    if len(argv) > 1:
        filein = argv[1]
    else:
        print("please run like 'python BullHolders.py [file]'")
        exit(1)
    timepoint = '2024-06-30'
    app = BullHolding(filein,timepoint)
    app.run()  
if __name__ == '__main__':
    fire.Fire(main)
    
    
    
    
    
  
    
    
    
    