import pandas as pd
import os
import math
import warnings
import xlrd

warnings.filterwarnings('ignore')
os.environ['TF_CPP_MIN_LOG_LEVEL']='3'

def sort_Groupby(Input_pd,col):
    grouped_pd = Input_pd.groupby(col)
    i = 0
    for item, item_df in grouped_pd:
        if i == 0:
            Result_pd = item_df
        else:
            Result_pd = Result_pd.append(item_df, ignore_index=True)
        i = i + 1
        #print(i)
    return Result_pd


####### 导入库存信息
# The file name of stock information
StockInfor_Filename = './Purchase_Rawdata/库存记录/ERP库存表格20210423.XLS'
pd_StockInfor1 = pd.DataFrame(pd.read_excel(StockInfor_Filename))
pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor1,'存货编码')
pd_StockInfor2 = pd_StockInfor1
#StockInfor_pd_byERPNUM = sort_Col_item(StockInfor_pd,2)
#pd_StockInfor_byERPNUM = sort_Groupby(pd_StockInfor,'存货编码')
#StockInfor_pd_byERPNUM.to_excel('./Purchase_Rawdata/results/ERPStock_byERPNUM_20210412.xlsx')

pd_StockInfor2['备注'] = None
BOM_folderNameStr = './BOM/'
i =0
for root, dirs, files in os.walk(BOM_folderNameStr):
    if i == 0:
        BOM_root = root
        BOM_dir_list = pd.Series(dirs)
    i = i+1

Batch_size = 1000
N = math.ceil(pd_StockInfor2.shape[0]/Batch_size)

for k in range(0,N):
    if (k+1)*Batch_size < pd_StockInfor2.shape[0]:
        pd_StockInfor = pd_StockInfor2[k * Batch_size: (k +1)* Batch_size]
        M = Batch_size
    else:
        pd_StockInfor = pd_StockInfor2[k * Batch_size :pd_StockInfor2.shape[0]]
        M = pd_StockInfor.shape[0]-Batch_size*k

    for i in range(0, M):
        target_ERP = pd_StockInfor.loc[pd_StockInfor.index[i],'存货编码']
        txt = None
        print('****Checking ERP number: ', target_ERP)
        print(i+k*Batch_size,' in total: ', pd_StockInfor2.shape[0])
        for j in range(0,BOM_dir_list.shape[0]):
            BOM_subfolder = BOM_folderNameStr + BOM_dir_list[j] + '/'
            #print('=======Serching in product BOM: ', BOM_dir_list[j])
            for root, dirs, files in os.walk(BOM_subfolder):
                for file in files:
                    if os.path.splitext(file)[1].lower() == '.xls':
                        wb = xlrd.open_workbook(root+'/'+file,logfile=open(os.devnull, 'w'))
                        pd_BOM = pd.DataFrame(pd.read_excel(wb, engine='xlrd'))
                        #print('------------Serching in BOM file:',root+'/'+file )
                        #if target_ERP == 'A010001':
                            #print(target_ERP)
                        pd_Match = pd_BOM[pd_BOM['子件编码(cpscode)']== target_ERP]
                        if not pd_Match.empty:
                            num = pd_Match['基本用量分子(ipsquantity)']
                            if txt == None:
                                txt =  str(BOM_dir_list[j]) + '_' + os.path.splitext(file)[0] + '_(' + str(num.values) + ') /'
                            else:
                                txt = str(txt) + str(BOM_dir_list[j]) + '_' + os.path.splitext(file)[0] + '_(' + str(num.values) + ') /'
                            #print(txt)
        pd_StockInfor.loc[pd_StockInfor.index[i],'备注'] = txt

    pd_StockInfor.to_excel('./Results/'+'ERP库存表格20210423_用途分析_'+str(k)+'.xlsx')




print('OK')