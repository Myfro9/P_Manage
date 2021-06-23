import pandas as pd
import os
import datetime
import xlwings as xw

## 自动生成询价单，并生成历年采购信息表，用于询价参考

## 输入：缺料表excel，历年采购合同记录excel
## 输出：询价单excel，历年采购参考信息excel

def ERP_Category_filter(pd_Infor,ERP_Kwd):
    if ERP_Kwd != 'others':
        if '存货编号(cInvCode)' in pd_Infor.columns :
            Bool_index = pd_Infor.loc[:,'存货编号(cInvCode)'].str.contains(ERP_Kwd).fillna(False)
        elif 'ERP' in pd_Infor.columns :
            Bool_index = pd_Infor.loc[:, 'ERP'].str.contains(ERP_Kwd).fillna(False)
        result_target = pd_Infor[Bool_index]
        result_left = pd_Infor[~Bool_index]
    else:
        result_target = pd_Infor
        result_left = None
    return result_target, result_left

def Supply_histroy(pd_BOMList, pd_Contract_Infor,kWD):
    result_byERP = pd.DataFrame(columns=['编号', '供应商个数','生产计划单号', 'ERP', '子件名称', '供应商名称','供应商类别', '总采购量', '平均价格', '最新价格', '最近合同日期','本次需采购量'])
    for i in range(pd_BOMList.shape[0]):
        ERP = pd_BOMList.loc[pd_BOMList.index[i],'ERP']
        Num_needed = pd_BOMList.loc[pd_BOMList.index[i],'需新增采购量']
        Contract_TargetERP = pd_Contract_Infor[pd_Contract_Infor.loc[:,'存货编号(cInvCode)'] == ERP]
        Contract_TargetERP_bySuplier  = Contract_TargetERP.groupby('供应商(cVenAbbName)')
        result1 = pd.DataFrame(columns=['编号','供应商个数','生产计划单号', 'ERP', '子件名称', '供应商名称', '总采购量', '平均价格', '最新价格', '最近合同日期','本次需采购量'])
        result0 = pd.DataFrame([[0,0,'A','A','A','无供应商记录',kWD,1.,0.1,0.1,'2007-01-01',1]],columns=['编号','供应商个数','生产计划单号','ERP','子件名称','供应商名称','供应商类别','总采购量','平均价格','最新价格','最近合同日期','本次需采购量'])
        for Supplier, Supplier_df in Contract_TargetERP_bySuplier:
            Num = Supplier_df.loc[:,'数量(iQuantity)'].sum()
            Total_Price = Supplier_df.loc[:,'价税合计(iSum)'].sum()
            Avg_Price = float(Total_Price / Num)
            Supplier_df['日期(dPODate)'] = pd.to_datetime(Supplier_df['日期(dPODate)'])
            Supplier_df['日期(dPODate)'] = Supplier_df['日期(dPODate)'].dt.strftime('%Y-%m-%d')
            Supplier_df = Supplier_df.sort_values(by='日期(dPODate)',ascending=False)
            result0.iat[0, 0] = i
            result0.iat[0, 1] = Contract_TargetERP_bySuplier.ngroups
            result0.iat[0, 2] = pd_BOMList.loc[pd_BOMList.index[i],'生产计划单号']
            result0.iat[0, 3] = ERP
            result0.iat[0, 4] = pd_BOMList.loc[pd_BOMList.index[i],'名称']
            result0.iat[0,5] = Supplier
            result0.iat[0, 6] = kWD
            result0.iat[0,7] = Num
            result0.iat[0,8] = Avg_Price
            result0.iat[0,9] = Supplier_df.loc[Supplier_df.index[0],'原币含税单价(iTaxPrice)']
            result0.iat[0,10] = Supplier_df.loc[Supplier_df.index[0],'日期(dPODate)']
            result0.iat[0,11] = Num_needed
            result1 = pd.concat([result1,result0])
        result1 = result1.sort_values(by='最近合同日期', ascending=False)
        if result1.shape[0] >0:
            result1.iloc[1::,0] =None
            result1.iloc[1::, 1] = None
            result1.iloc[1::, 2] = None
            result1.iloc[1::, 3] = None
            result1.iloc[1::, 11] = None

        result_byERP = pd.concat([result_byERP,result1])
        result_byERP['目标供应商'] = None
        result_byERP['目标价格'] = None
        result_byERP['目标采购数量'] = None
    #result_byERP.to_excel(FolderNameStr + FileNameStr_Target + '_' + kWD +'_Suppliers.xlsx')
    return result_byERP

def Ask4Price(pd_BOMList_need2Buy, kWD):
    pd_Ask4Price = pd_BOMList_need2Buy[['生产计划单号','ERP', '名称',  '需新增采购量']]
    pd_Ask4Price.columns = ['生产计划单号','ERP编码', '型号', '需求数量']
    pd_Ask4Price['供应商类别'] = kWD
    pd_Ask4Price['报价数量'] = None
    pd_Ask4Price['报价单价（元）'] = None
    pd_Ask4Price['交货期(天）'] = None
    pd_Ask4Price.index = range(pd_Ask4Price.shape[0])
    #pd_Ask4Price_IC.to_excel(FolderNameStr + FileNameStr_Target + '_' + kWD + '_Ask4quote.xlsx')
    return pd_Ask4Price


####### 导入所有的采购合同
FolderNameStr = './Purchase_Rawdata/21采购合同记录/'
FileNameStr = 'ERP20150101-20210420.xls'
pd_Contract_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr))

####### 导入需要采购的BOM列表
FileNameStr = 'YF202106-B.xlsx'
FolderNameStr = './Results/'
FolderNameStr1 = './Results/YF202106-B_询价表/'
FileNameStr_Target = 'YF202106-B'
pd_BOMList_need2Buy = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr))

#######
####### 从采购合同中过滤所有芯片采购类合同信息，'M0101'
pd_Contract_Infor_IC,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0101')
## 从采购BOM列表中过滤所有芯片采购类需求
pd_ICList_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0101')
## 从芯片采购列表中，逐个分析历史采购合同信息，生成芯片类历史参考价格，以及询价单
pd_Supply_histroy_IC = Supply_histroy(pd_ICList_need2Buy, pd_Contract_Infor_IC,  'IC')
pd_Ask4Price_IC = Ask4Price(pd_ICList_need2Buy, 'IC')



#######
####### 从采购合同中过滤所有PCB板采购类合同信息，'M0106'
pd_Contract_Infor_PCB,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0106')
## 从采购BOM列表中过滤所有PCB采购类需求
pd_PCBList_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0106')
## 从PCB采购列表中，逐个分析历史采购合同信息，生成PCB类历史参考价格，以及询价单
pd_Supply_histroy_PCB = Supply_histroy(pd_PCBList_need2Buy, pd_Contract_Infor_PCB,  'PCB')
pd_Ask4Price_PCB = Ask4Price(pd_PCBList_need2Buy, 'PCB')

#######
####### 从采购合同中过滤所有RLC采购类合同信息，'M0102','M0103','M0104','M0105'
pd_Contract_Infor_R,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0102')
pd_Contract_Infor_L,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0104')
pd_Contract_Infor_L2,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0105')
pd_Contract_Infor_C,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0103')
pd_Contract_Infor_RLC = pd.concat([pd_Contract_Infor_R,pd_Contract_Infor_L,pd_Contract_Infor_C,pd_Contract_Infor_L2])
## 从采购BOM列表中过滤所有PCB采购类需求
pd_R_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0102')
pd_L_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0104')
pd_L2_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0105')
pd_C_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0103')
pd_RLC_List_need2Buy = pd.concat([pd_R_List_need2Buy,pd_L_List_need2Buy,pd_C_List_need2Buy,pd_L2_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成RLC类历史参考价格，以及询价单
pd_Supply_histroy_RLC = Supply_histroy(pd_RLC_List_need2Buy, pd_Contract_Infor_RLC,  'RLC')
pd_Ask4Price_RLC = Ask4Price(pd_RLC_List_need2Buy, 'RLC')

#######
####### 从采购合同中过滤所有晶振及其它采购类合同信息，'M0108','M0107'
pd_Contract_Infor_OC,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0108')
pd_Contract_Infor_OC_misc,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0107')
pd_Contract_Infor_OCs = pd.concat([pd_Contract_Infor_OC,pd_Contract_Infor_OC_misc])
## 从采购BOM列表中过滤所有晶振及其它采购类需求
pd_OC_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0108')
pd_OC_misc_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0107')
pd_OCs_List_need2Buy = pd.concat([pd_OC_List_need2Buy,pd_OC_misc_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成晶振及其它类历史参考价格，以及询价单
pd_Supply_histroy_OCs = Supply_histroy(pd_OCs_List_need2Buy, pd_Contract_Infor_OCs,  'OCs')
pd_Ask4Price_OCs = Ask4Price(pd_OCs_List_need2Buy, 'OCs')

#######
####### 从采购合同中过滤所有'电源-风扇-LED'采购类合同信息，'M0201','M0202','M0204','M0205'
pd_Contract_Infor_Power,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0201')
pd_Contract_Infor_Fan,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0202')
pd_Contract_Infor_LED,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0204')
pd_Contract_Infor_Display,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0205')
pd_Contract_Infor_Power_Fan_LED_Display = pd.concat([pd_Contract_Infor_Power,pd_Contract_Infor_Fan,pd_Contract_Infor_LED,pd_Contract_Infor_Display])
## 从采购BOM列表中过滤所有'电源-风扇-LED'采购类需求
pd_Power_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0201')
pd_Fan_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0202')
pd_LED_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0204')
pd_Display_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0205')
pd_Power_Fan_LED_Display_List_need2Buy = pd.concat([pd_Power_List_need2Buy,pd_Fan_List_need2Buy,pd_LED_List_need2Buy,pd_Display_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成'电源-风扇-LED'类历史参考价格，以及询价单
pd_Supply_histroy_Power_Fan_LED_Display = Supply_histroy(pd_Power_Fan_LED_Display_List_need2Buy, pd_Contract_Infor_Power_Fan_LED_Display,  'Power_Fan_LED_Display')
pd_Ask4Price_Power_Fan_LED_Display = Ask4Price(pd_Power_Fan_LED_Display_List_need2Buy, 'Power_Fan_LED_Display')

#######
####### 从采购合同中过滤所有'结构件'采购类合同信息，'M0301','M0302','M0304','M0306'
pd_Contract_Infor_mechanic0,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0301')
pd_Contract_Infor_mechanic1,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0302')
pd_Contract_Infor_mechanic2,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0304')
pd_Contract_Infor_mechanic3,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0306')
pd_Contract_Infor_mechanic = pd.concat([pd_Contract_Infor_mechanic0,pd_Contract_Infor_mechanic1,pd_Contract_Infor_mechanic2,pd_Contract_Infor_mechanic3])
## 从采购BOM列表中过滤所有'结构件'采购类需求
pd_mechanic0_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0301')
pd_mechanic1_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0302')
pd_mechanic2_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0304')
pd_mechanic3_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0306')
pd_mechanic_List_need2Buy = pd.concat([pd_mechanic0_List_need2Buy,pd_mechanic1_List_need2Buy,pd_mechanic2_List_need2Buy,pd_mechanic3_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成'结构件'类历史参考价格，以及询价单
pd_Supply_histroy_mechanic =Supply_histroy(pd_mechanic_List_need2Buy, pd_Contract_Infor_mechanic,  'mechanic')
pd_Ask4Price_mechanic = Ask4Price(pd_mechanic_List_need2Buy,'mechanic')


#######
####### 从采购合同中过滤所有'螺丝及其它金属件'采购类合同信息，'M0303','M0305'
pd_Contract_Infor_Screw,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0303')
pd_Contract_Infor_Screw_misc0,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0305')
pd_Contract_Infor_Screw_misc = pd.concat([pd_Contract_Infor_Screw,pd_Contract_Infor_Screw_misc0])
## 从采购BOM列表中过滤所有'螺丝及其它金属件'采购类需求
pd_Screw_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0303')
pd_Screw_misc0_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0305')
pd_Screw_misc_List_need2Buy = pd.concat([pd_Screw_List_need2Buy,pd_Screw_misc0_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成'螺丝及其它金属件'类历史参考价格，以及询价单
pd_Supply_histroy_Screw_misc = Supply_histroy(pd_Screw_misc_List_need2Buy, pd_Contract_Infor_Screw_misc,  'Screw_misc')
pd_Ask4Price_Screw_misc = Ask4Price(pd_Screw_misc_List_need2Buy, 'Screw_misc')

#######
####### 从采购合同中过滤所有'连接件'采购类合同信息，'M0203','M0206','M0207'
pd_Contract_Infor_Connect0,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0203')
pd_Contract_Infor_Connect1,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0206')
pd_Contract_Infor_Connect2,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M0207')
pd_Contract_Infor_Connect = pd.concat([pd_Contract_Infor_Connect0,pd_Contract_Infor_Connect1,pd_Contract_Infor_Connect2])
## 从采购BOM列表中过滤所有'连接件'采购类需求
pd_Connect0_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0203')
pd_Connect1_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0206')
pd_Connect2_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M0207')
pd_Connect_List_need2Buy = pd.concat([pd_Connect0_List_need2Buy,pd_Connect1_List_need2Buy,pd_Connect2_List_need2Buy])
## 从PCB采购列表中，逐个分析历史采购合同信息，生成'连接件'类历史参考价格，以及询价单
pd_Supply_histroy_Connect = Supply_histroy(pd_Connect_List_need2Buy, pd_Contract_Infor_Connect,  'Connect')
pd_Ask4Price_Connect = Ask4Price(pd_Connect_List_need2Buy, 'Connect')

#######
####### 从采购合同中过滤所有'射频件'采购类合同信息，'M040'
pd_Contract_Infor_RF,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'M040')
## 从采购BOM列表中过滤所有PCB采购类需求
pd_RFList_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'M040')
## 从PCB采购列表中，逐个分析历史采购合同信息，生成PCB类历史参考价格，以及询价单
pd_Supply_histroy_RF = Supply_histroy(pd_RFList_need2Buy, pd_Contract_Infor_RF,  'RF')
pd_Ask4Price_RF =  Ask4Price(pd_RFList_need2Buy, 'RF')

#######
####### 从采购合同中过滤所有'焊接'采购类合同信息，'H0'
pd_Contract_Infor_Soldering,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'H01')
## 从采购BOM列表中过滤所有PCB采购类需求
pd_Soldering_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'H01')
## 从PCB采购列表中，逐个分析历史采购合同信息，生成PCB类历史参考价格，以及询价单
pd_Supply_histroy_Soldering = Supply_histroy(pd_Soldering_List_need2Buy, pd_Contract_Infor_Soldering,  'Soldering')
pd_Ask4Price_Soldering = Ask4Price(pd_Soldering_List_need2Buy, 'Soldering')

#######
####### 从采购合同中过滤所有'其它'采购类合同信息，
pd_Contract_Infor_others,pd_Contract_Infor = ERP_Category_filter(pd_Contract_Infor,'others')
## 从采购BOM列表中过滤所有PCB采购类需求
pd_others_List_need2Buy,pd_BOMList_need2Buy = ERP_Category_filter(pd_BOMList_need2Buy,'others')
## 从PCB采购列表中，逐个分析历史采购合同信息，生成PCB类历史参考价格，以及询价单
pd_Supply_histroy_others = Supply_histroy(pd_others_List_need2Buy, pd_Contract_Infor_others,  'others')
pd_Ask4Price_others = Ask4Price(pd_others_List_need2Buy, 'others')

pd_Supply_histroy = pd.concat([pd_Supply_histroy_IC, pd_Supply_histroy_PCB, pd_Supply_histroy_RLC, pd_Supply_histroy_OCs, pd_Supply_histroy_Power_Fan_LED_Display,pd_Supply_histroy_mechanic,pd_Supply_histroy_Screw_misc,pd_Supply_histroy_Connect,pd_Supply_histroy_RF,pd_Supply_histroy_Soldering,pd_Supply_histroy_others])
pd_Ask4Price = pd.concat([pd_Ask4Price_IC,pd_Ask4Price_PCB,pd_Ask4Price_RLC,pd_Ask4Price_OCs,pd_Ask4Price_Power_Fan_LED_Display,pd_Ask4Price_mechanic,pd_Ask4Price_Screw_misc,pd_Ask4Price_Connect,pd_Ask4Price_RF,pd_Ask4Price_Soldering,pd_Ask4Price_others])

if ~os.path.exists(FolderNameStr1):
    os.makedirs(FolderNameStr1)
pd_Supply_histroy.to_excel(FolderNameStr1+FileNameStr_Target+'_供应商信息参考.xlsx')
pd_Ask4Price.to_excel(FolderNameStr1+FileNameStr_Target+'_采购询价表.xlsx')

print('test')
'''FolderNameStr = './Purchase_Rawdata/'
FileNameStr0 = 'Contract Template.xlsx'
FileNameStr1 = 'Contract target.xlsx'

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False  # 是否实时刷新excel程序的显示内容
wb = app.books.open(FolderNameStr + FileNameStr0)
ws = wb.sheets[0]
print(ws.name)

ws.range('A1').value = '买卖合同test1'

wb.save(FolderNameStr + FileNameStr1)
wb.close()
app.quit()

'''