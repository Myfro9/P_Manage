import pandas as pd
import os
import time
import rd_DataBase

supplier_list = ['上海汇甬电子有限公司','江苏德是和通信科技有限公司','斯伏电源系统（上海）有限公司',\
                 '太仓超建精密五金有限公司','苏州祥跃迈腾电子有限公司','广州市新原空气净化工程有限公司',\
                '广州锋毅包装材料制品有限公司','上海诗吟实业有限公司','上海得格实业有限公司','上海爱力特电子有限公司',\
                 '上海临胤实业有限公司','上海一鹏电子科技有限公司','昆山金屹敏电子有限公司']
pd_Task = pd.DataFrame({'Task':['RW202112-G-1','RW202112-G-2','RW202112-H-1','RW202112-H-2','RW202112-H-3',\
                                'RW202201-A-1','RW202201-A-2','RW202201-A-3','RW202201-A-4','RW202201-A-5',\
                                'RW202201-A-6','RW202201-A-7','RW202201-A-8','RW202201-A-9',\
                                'RW202201-B-1','RW202201-B-2','RW202201-B-3','RW202201-B-4',\
                                'RW202201-C-1','RW202201-C-2','RW202201-C-3','RW202201-C-4','RW202201-C-5',\
                                'RW202201-C-6','RW202201-C-7','RW202201-D',] })
FolderNameStr_quote = './Purchase_Rawdata/23询价结果/'
FileNameStr_quote = '合并汇总表-版本76-20220420.xlsx'
pd_QuoteInfo = pd.DataFrame(pd.read_excel(FolderNameStr_quote+FileNameStr_quote,sheet_name='操作'))

pd_VirtualStock,pd_PrjList,FileNameStr_PrjList,StockInfor_Filename,FileNameStr_contract,\
           Contract_FileNameStr1,Contract_FileNameStr2,FileNameStr_outstock,FileNameStr_instock,\
           FileNameStr_outsourcing,FileNameStr_outsourcingBack,\
            pd_StockInfor,pd_Contract_inTask,pd_Outstock_inTask,pd_Instock_inTask,pd_OutSourcing_inTask,pd_OutSourcingBack_inTask = rd_DataBase.Purchase_Rawdata_analyze(None)

#### 导入采购到货记录（张永亮）  --- I （注意，不是采购入库记录，到货与入库之间有时间差！），后面可以滤出在产任务单的到货记录
FolderNameStr = './Purchase_Rawdata/24采购到货记录/'
FileNameStr = '2021ERP到货单列表-20220210.xls'   # 2021 年记录
FileNameStr_1 = '2022ERP到货单列表-20220504.xls' # 2022 年记录
pd_PurchaseArrive_0 = pd.DataFrame(pd.read_excel(FolderNameStr + FileNameStr))
pd_PurchaseArrive_1 = pd.DataFrame(pd.read_excel(FolderNameStr + FileNameStr_1))
pd_PurchaseArrive = pd.concat([pd_PurchaseArrive_1,pd_PurchaseArrive_0])
pd_PurchaseArrive_inTask = rd_DataBase.OnGoing_Prjinfo_Filter(pd_PrjList, pd_PurchaseArrive, 0)

####  导入ERP材料入库记录（许金鲍） -- II ，并整理出对应在产任务单的材料入库记录
pd_ERPInstock_inTask = pd_Instock_inTask

####  导入ERP材料出库记录（许金鲍） -- III ，并整理出对应在产任务单的材料出库记录
pd_ERPOutstock_inTask = pd_Outstock_inTask

####  导入委外材料出库记录（许金鲍） --- IV
pd_ERPOutsourcing_inTask = pd_OutSourcing_inTask

TimeStr = time.strftime("%Y-%m-%d-%H", time.localtime(time.time()))
Result_Folder = './Result/特殊供应商合同及供货情况_' + TimeStr + '/'
if not os.path.exists(Result_Folder):
    os.makedirs(Result_Folder)

for i in range(len(supplier_list)):
    pd_match0 = pd_QuoteInfo[pd_QuoteInfo['渠道']==supplier_list[i]]
    pd_match1 = pd_match0[pd_match0['生产计划单号'].isin(pd_Task['Task'])]
    assert not pd_match1.empty, "没有找到{}的销售记录".format(supplier_list[i])
    pd_result = pd_match1.loc[:,['生产计划单号','ERP编码','型号','报价数量','单价（元）','小计','渠道','盖章日期','规格']]
    pd_result['到货数量'] = None
    pd_result['入库数量']  = None
    pd_result['付款审批'] = None
    for ii in range(pd_result.shape[0]):
        Task = pd_result.loc[pd_result.index[ii],'生产计划单号']
        ERP = pd_result.loc[pd_result.index[ii],'ERP编码']
        result0 = pd_PurchaseArrive_inTask.loc[(pd_PurchaseArrive_inTask['存货编码(cinvcode)']==ERP) & \
                                               (pd_PurchaseArrive_inTask['备注(cmemo)'].str.contains(Task)) & \
                                               (pd_PurchaseArrive_inTask['供应商(cvenabbname)'] == supplier_list[i])]
        if not result0.empty:
            num = result0['主计量数量(iquantity)'].sum()
            pd_result.loc[pd_result.index[ii], '到货数量'] = num
        result1 = pd_ERPInstock_inTask.loc[ (pd_ERPInstock_inTask['存货编码(cInvCode)'] ==ERP) & \
                                            (pd_ERPInstock_inTask['备注(cMemo)'].str.contains(Task)) & \
                                            (pd_ERPInstock_inTask['供货单位(cVenAbbName)'] == supplier_list[i])]
        if not result1.empty:
            num = result1['数量(iQuantity)'].sum()
            pd_result.loc[pd_result.index[ii], '入库数量'] = num
    pd_result.to_excel(Result_Folder +supplier_list[i] +'.xlsx')
    print(supplier_list[i])