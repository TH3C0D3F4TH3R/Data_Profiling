import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import xlsxwriter
from io import BytesIO

SMALL_SIZE = 7
plt.rc('font',size = SMALL_SIZE)
plt.rc('axes',titlesize = SMALL_SIZE)

def data_import(path):
    data =pd.read_csv(path)
    return data

def data_prep(df):
    if 'target' in df.columns.str.lower():
        i=np.where(df.columns.str.lower()=='target')[0][0]
        test= df.copy()
        test['accept/decline'] = ['accept' if x==1 else 'decline' for x in test[test.columns[i]]]
        test['accept'] = [1 if x==1 else 0 for x in test[test.columns[i]]]
        test['decline'] = [1 if x==0 else 0 for x in test[test.columns[i]]]
        return test
    else:
        print("No target found")
        return

def var_sel(df):
    var_list=[]
    for i in range(0,len(df.columns)):
        print(i,df.columns[i])
    numbers = [int(n) for n in input('Input the variables index for profiling, separated by commas: ').split(',')]
    
    for j in range(0,len(numbers)):
        var_list.append(df.columns[numbers[j]])
        
    return var_list

def pivot_col(df,var_list):
    table2 = pd.DataFrame()
    table2.name = 'accept/decline'
    table2.reindex(columns=['field','lvl','accept','decline'])
    for x in df[var_list].columns:
        table1 = pd.pivot_table(df,index = x,values = 'target',columns = 'accept/decline',aggfunc = 'count')
        table1['field'] = x
        table1.reset_index(inplace=True)
        table1.reindex(columns = ['field',x,'accept','decline'])
        table1 = table1.rename(columns = {x:'lvl'})
        table2 = pd.concat([table2,table1],axis=0)
    table3 = table2.reindex(columns = ['field','lvl','accept','decline'])
    return table3

def make_table(df,var):
    temp = df.loc[df['field']==var]
    temp_name = temp['field'][0]
    temp1 = temp.drop('field',axis=1)
    temp1.fillna(0,inplace=True)
    temp1['Claims'] = [x+y for x,y in zip(temp1['accept'],temp1['decline'])]
    temp1['Claim Distribution %'] = [(x/temp1['Claims'].sum())*100 for x in temp1['Claims']]
    temp1['Accept Rate %'] = [(x/y)*100 for x,y in zip(temp1['accept'],temp1['Claims'])]
    temp1['Decline Rate %'] = [(x/y)*100 for x,y in zip(temp1['decline'],temp1['Claims'])]
    
    temp1 = temp1.sort_values(by = 'lvl')
    if temp1['lvl'].dtype in ['int64','float64']:
        temp1['lvl'] = temp1['lvl'].astype('str')
    temp2 = temp1.copy()
    temp2 = pd.DataFrame(temp2)
    return temp2

def plot_graph(df,var):
    FONT_SIZE=8
    AXES_SIZE = 9
    plt.rc('font',weight = 'bold')
    plt.rcParams['axes.labelweight']='bold'
    plt.rc('font',size = FONT_SIZE)
    plt.rc('axes',titlesize = AXES_SIZE)
    
    fig=plt.figure(figsize=[10,5])
    ax1 = fig.add_subplot(111)
    ax1.bar(x=df['lvl'],height = df['Claim Distribution %'],label = 'Claim Distribution %',edgecolor = 'black')
    ax1.set_ylabel('Claim Distribution %')
    ax1.legend(loc='upper right',bbox_to_anchor = (-0.15,0.90),fontsize ='medium')
    ax1.set_xlabel(var)
    ax1.set_xticklabels(df['lvl'],rotation = 90)
    ax2 = ax1.twinx()
    
    ax2.plot(df['lvl'],df['Accept Rate %'], 'r-',marker='o')
    ax2.set_ylabel('Accept Rate %')
    ax2.legend(loc='upper right',bbox_to_anchor = (-0.15,0.97),fontsize ='medium')
    
    return fig

def data_profile(df_ip,var_list):
    
    cwd = os.getcwd()
    path = cwd + "/data_profile.xlsx"
    writer = pd.ExcelWriter(path,engine='xlsxwriter')
    workbook = writer.book
    
    d = {'col1':[1],'col2':[2]}
    df = pd.DataFrame(data=d)
    
    for i in range(0,len(var_list)):
        
        df.to_excel(writer,sheet_name =var_list[i])
        
    table = pivot_col(df_ip,var_list)
    
    for i in range(0,len(var_list)):
        temp = pd.DataFrame()
        imgdata = BytesIO()
        nm = var_list[i]
        
        temp = make_table(table,nm)
        temp = pd.DataFrame(temp)
        temp.to_excel(writer,sheet_name = nm,index=False)
        
        
        fig = plot_graph(temp,nm)
        workbook = writer.book
        worksheet = writer.sheets[var_list[i]]
        
        header_format = workbook.add_format({'bottom':2,'bg_color':'#F9DA04'})
        
        for col_num,value in enumerate(temp.columns.values):
            worksheet.write(0,col_num,value,header_format)
            
        format1 = workbook.add_format()
        format1.set_bottom(7)
        
        format2 = workbook.add_format()
        format2.set_top(2)
        format2.set_bottom(7)
        
        worksheet.conditional_format(1,0,1,temp.shape[1]-1,
                                    {'type':'cell',
                                    'criteria':'<>',
                                    'value':'""',
                                    'format':format2})
        
        worksheet.conditional_format(1,0,temp.shape[0],temp.shape[1]-1,
                                    {'type':'cell',
                                    'criteria':'<>',
                                    'value':'""',
                                    'format':format1})
        
        fig.savefig(imgdata,quality=95,format='png',pd_inches=0.3,bbox_inches='tight')
        
        imgdata.seek(0)
        
        worksheet.insert_image(1,9,"",{'image_data':imgdata})
        
        plt.close()
    writer.save()
    imgdata.truncate()
    
    print('Data Profiling has been completed! You can access the output workbook at {path}'.format(path=path))
        
        
        