import pandas as pd
import os
import Purchase_Production_ListGen


Prj = 'RW202201-C-7'


# 针对BOM更新，对比更新前后的Purchase_BOM,列举出需要取消采购的部分和需要新增采购的部分
log_FileNameStr = './PurchaseBOM_询价表_log.xlsx'
pd_log = pd.DataFrame(pd.read_excel(log_FileNameStr))
target = pd_log[pd_log['生产计划单']==Prj]
Old_PurchaseBOM = target['PurchaseBOM'].values[0]
Old_PurchaseBOM_FolderNameStr = './Results/'+ Old_PurchaseBOM + '/'
log_file2 = Old_PurchaseBOM_FolderNameStr + 'PurchaseBOM_inforlog.txt'
for line in open(log_file2):
    if '在产任务单列表' in line:
        FileNameStr_PrjList = line.strip('在产任务单列表').strip('：').strip('\n')
    elif '库存列表' in line:
        StockInfor_Filename = line.strip('库存列表').strip('：').strip('\n')
    elif '历史采购合同列表' in line:
        FileNameStr_contract = line.strip('历史采购合同列表').strip('：').strip('\n')
    elif '新增采购合同列表1' in line:
        Contract_FileNameStr1 = line.strip('新增采购合同列表1').strip('：').strip('\n')
    elif '出库信息' in line:
        FileNameStr_outstock = line.strip('出库信息').strip('：').strip('\n')
    elif '入库信息' in line:
        FileNameStr_instock = line.strip('入库信息').strip('：').strip('\n')
    elif '委外信息' in line:
        FileNameStr_outsourcing = line.strip('委外信息').strip('：').strip('\n')
    elif '委外返回信息' in line:
        FileNameStr_outsourcingBack = line.strip('委外返回信息').strip('：').strip('\n')

argv_list = [1,FileNameStr_PrjList,StockInfor_Filename,FileNameStr_contract,Contract_FileNameStr1,FileNameStr_outstock,FileNameStr_instock,FileNameStr_outsourcing,FileNameStr_outsourcingBack]
New_PurchaseBOM_FolderNameStr = Purchase_Production_ListGen.main(argv_list)

print('ok')