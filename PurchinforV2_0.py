import pandas as pd


# The inputs file names:
FolderNameStr = './Purchase_Rawdata/21采购合同记录/'
FileNameStr1 = 'ERP20150101-20210420.xls'
'''FileNameStr1 = 'ERP2016.xls'
FileNameStr2 = 'ERP2017.xls'
FileNameStr3 = 'ERP2018.xls'
FileNameStr4 = 'ERP2019.xls'
FileNameStr5 = 'ERP20200101-20210420.xls'''''

# The outputs file names:
Results_FolderNameStr = './Results/'
FileNameStr_result_byERPnum = 'Purchase_byERPnum_202104.xlsx'
FileNameStr_result_bySpplier = 'Purchase_bySupplier_202011.xlsx'
FileNameStr_SpplierSumPrice = 'SupplierSumPrice_202011.xlsx'
FileNameStr_result_RefPrice_byERPnum0 = 'RefPrice_byERPnum_202104.xlsx'
FileNameStr_result_RefPrice_byERPnum = 'RefPrice_byERPnum.xlsx'

xls = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr1))

# 近期询价结果的数据库文件
FolderNameStr_new = './Purchase_Rawdata/23询价结果/'
FileNameStr2 = '汇总YF06AB_GC06AB_RW06DEF12G12H1-23-20210902.xlsx'   # 近期询价结果
FileNameStr3 = '汇总表GC8A-RW8B2-8C-3-20210902.xlsx'                 # 近期询价结果
### 合并近期询价结果的数据库
xls1 = pd.DataFrame(pd.read_excel(FolderNameStr_new+FileNameStr2,sheet_name='操作'))
xls2 = pd.DataFrame(pd.read_excel(FolderNameStr_new+FileNameStr3,sheet_name='操作'))
xls1 = xls1[xls1['报价数量']>0]
xls1 = xls1[xls1['小计']>0]
xls2 = xls2[xls2['报价数量']>0]
xls2 = xls2[xls2['小计']>0]
xls_new = pd.concat([xls1,xls2])
xls_new = xls_new.loc[:,['ERP编码','型号','报价数量','小计','单价（元）']]
xls_new.columns = ['存货编号(cInvCode)','存货名称(cInvName)','数量(iQuantity)','价税合计(iSum)','原币含税单价(iTaxPrice)']


'''xls2 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr2))
xls3 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr3))
xls4 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr4))
xls5 = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr5))'''

#xls = pd.concat([xls1,xls2,xls3,xls4,xls5])

#result = xls.iloc[:, [1,2,3,4,5,6,7,8,10,12,13]]
if False : # Re-order the contract record in ERPnum sequence
    Item_counts = xls.iloc[:, 5].value_counts()
    counts_result = pd.DataFrame({'col1': Item_counts.index, 'col2': Item_counts.values})
    target = counts_result.iloc[0,0]
    result_byERPnum= xls.loc[xls['存货编号(cInvCode)']==target]
    #for i in range(1,3):
    for i in range(1,counts_result.shape[0]):
        target = counts_result.iloc[i,0]
        result1 = xls.loc[xls['存货编号(cInvCode)']==target]
        result_byERPnum = pd.concat([result_byERPnum,result1])
        print(i)
    # result_byERPnum.drop_duplicates('订单编号(cPOID)','first',inplace=True)
    result_byERPnum.to_excel(Results_FolderNameStr+FileNameStr_result_byERPnum)

if False :  # Re-order the contract record in supplier sequence
    Supplier_counts = xls.iloc[:, 3].value_counts()
    Supplier_counts_result = pd.DataFrame({'col1': Supplier_counts.index, 'col2': Supplier_counts.values})
    target = Supplier_counts_result.iloc[0,0]
    result_bySupplier= xls.loc[xls['供应商(cVenAbbName)']==target]
    #for i in range(1,3):
    for i in range(1,Supplier_counts_result.shape[0]):
        target = Supplier_counts_result.iloc[i,0]
        result1 = xls.loc[xls['供应商(cVenAbbName)']==target]
        result_bySupplier = pd.concat([result_bySupplier,result1])
        print('Supplier:',i)
    # result_bySupplier.drop_duplicates('订单编号(cPOID)','first',inplace=True)
    result_bySupplier.to_excel(Results_FolderNameStr+FileNameStr_result_bySpplier)


if False: # Calc total contract price for each supplier
    xls_bySupplier = pd.DataFrame(pd.read_excel(Results_FolderNameStr+FileNameStr_result_bySpplier))
    Supplier_sr = xls_bySupplier.iloc[:, 4].value_counts()
    Supplier_pd = pd.DataFrame({'Name': Supplier_sr.index, 'Total Contracts': Supplier_sr.values})
    target = Supplier_pd.iloc[0,0]
    targetSupplier_pd= xls_bySupplier.loc[xls_bySupplier['供应商(cVenAbbName)']==target]
    SumPrice = targetSupplier_pd['价税合计(iSum)'].sum()
    Supplier_pd['Total Price'] = 0
    Supplier_pd.iloc[0,2]=SumPrice
    # for i in range(1,3):
    for i in range(1,Supplier_pd.shape[0]):
        target = Supplier_pd.iloc[i,0]
        targetSupplier_pd= xls_bySupplier.loc[xls_bySupplier['供应商(cVenAbbName)']==target]
        SumPrice = targetSupplier_pd['价税合计(iSum)'].sum()
        Supplier_pd.iloc[i, 2] = SumPrice
        print('Supplier_sumPrice:',i)
    Supplier_pd.sort_values(by="Total Price",ascending=False,inplace=True)
    Supplier_pd.to_excel(Results_FolderNameStr+FileNameStr_SpplierSumPrice)

def Price_Analy_byERPnum(ERP_pd):

    Total_ItemNum = ERP_pd['数量(iQuantity)'].sum()
    Total_price = ERP_pd['价税合计(iSum)'].sum()
    AvgPrice_byItemNum = Total_price/Total_ItemNum
    Lowest_Price = min(ERP_pd['原币含税单价(iTaxPrice)'])
    ERP_pd['Var']= ERP_pd['原币含税单价(iTaxPrice)'].map(lambda x: round((x-AvgPrice_byItemNum)/AvgPrice_byItemNum,2))
    list_Var_sr = ERP_pd['Var'].value_counts()
    Var_pd = pd.DataFrame({'Var':list_Var_sr.index,'Contract Qt':list_Var_sr.values})
    Var_pd['Qt'] = 0
    Var_pd['Unit_Price'] = 0
    Var_pd['Percent'] = 0
    for i in range(0,Var_pd.shape[0]):
        target = Var_pd.iloc[i,0]
        target_pd = ERP_pd.loc[ERP_pd['Var'] == target]
        Var_qt = target_pd['数量(iQuantity)'].sum()
        Var_Price = target_pd['价税合计(iSum)'].sum() / Var_qt
        Var_Percent = Var_qt / Total_ItemNum
        Var_pd.iloc[i, 2] = Var_qt
        Var_pd.iloc[i, 3] = Var_Price
        Var_pd.iloc[i, 4] = Var_Percent

    Ref_pd = Var_pd[((Var_pd['Var'] <= 0.1) & (Var_pd['Var'] >= -0.1)) | (Var_pd['Percent'] > 0.1) ]
    # Ref_pd['Total_Price'] = Ref_pd.apply(lambda x : x['Unit_Price'] * x['Qt'],axis=1)
    Ref_pd.eval('Total_Price = Unit_Price * Qt', inplace=True)
    Bool = ((Var_pd['Var']>0.5) & (Var_pd['Percent'] > 0.1) )
    Warning_Code =0
    if True in Bool.values:
        print('Warning_Price#1: There is too high contract price for ERP number ', ERP_pd.iloc[0,0])
        Warning_Code = 1

    Bool = ((ERP_pd['Var']<-0.5)& (Var_pd['Percent'] > 0.1) )
    if True in Bool.values :
        print('Warning_Price#1: There is too low contract price for ERP number ', ERP_pd.iloc[0,0])
        Warning_Code = 2

    Total_ItemNum = Ref_pd['Qt'].sum()
    Total_price = Ref_pd['Total_Price'].sum()
    AvgPrice_byItemNum = Total_price / Total_ItemNum
    return [AvgPrice_byItemNum,Lowest_Price,Warning_Code,Var_pd]


def price_analy(xls_byERPnum):
    list_sr = xls_byERPnum.loc[:, '存货编号(cInvCode)'].value_counts()
    ERP_Price_pd = pd.DataFrame({'ERP': list_sr.index, 'Contract Qt': list_sr.values})
    ERP_Price_pd['Ref_UnitPrice'] = 0
    ERP_Price_pd['Name'] = 'Name'
    ERP_Price_pd['Warning_Code'] = 0
    ERP_Price_pd['Lowest_Price'] = 0
    for i in range(0, ERP_Price_pd.shape[0]):
        #   for i in range(0,20):
        target = ERP_Price_pd.iloc[i, 0]
        target_pd = xls_byERPnum.loc[xls_byERPnum['存货编号(cInvCode)'] == target]
        [ref_Price, Lowest_Price, Warning_Code, Var_pd] = Price_Analy_byERPnum(target_pd)
        print(i, 'ref_Price is:', ref_Price)
        ERP_Price_pd.iloc[i, 2] = ref_Price
        ERP_Price_pd.iloc[i, 3] = target_pd.iloc[0, 1]
        ERP_Price_pd.iloc[i, 4] = Warning_Code
        ERP_Price_pd.iloc[i, 5] = Lowest_Price
    return ERP_Price_pd

if True:  ## Calculate the ref_price and lowest price according to contract records ordered in ERPnum
    if False:   ## 重新计算历史采购价格参考信息，因为计算时间比较长，所以存储成文件，不用每次再算
        xls_byERPnum = pd.DataFrame(pd.read_excel(Results_FolderNameStr+FileNameStr_result_byERPnum))  ## 导入历史采购记录数据库（也是在本程序中处理的结果）
        xls_byERPnum = xls_byERPnum.loc[:,['存货编号(cInvCode)','存货名称(cInvName)','数量(iQuantity)','价税合计(iSum)','原币含税单价(iTaxPrice)']]
        ERP_Price_pd_history = price_analy(xls_byERPnum)
        ERP_Price_pd_history.to_excel(Results_FolderNameStr + FileNameStr_result_RefPrice_byERPnum0)


    ERP_Price_pd_history = pd.DataFrame(pd.read_excel(Results_FolderNameStr + FileNameStr_result_RefPrice_byERPnum0))
    ERP_Price_pd_history.drop(columns=ERP_Price_pd_history.columns[0],axis=1,inplace=True)
    xls_byERPnum_new = xls_new   ## 导入近期的询价结果数据库
    ERP_Price_pd_new = price_analy(xls_byERPnum_new)

    ERP_Price_pd_history['history Price'] =0
    ERP_Price_pd_new['history Price'] = 0
    for i in range(0,ERP_Price_pd_history.shape[0]):
        ERP = ERP_Price_pd_history.iloc[i,0]
        target = ERP_Price_pd_new[ERP_Price_pd_new['ERP'].str.contains(ERP).fillna(False)]
        if not target.empty:
            target['history Price'] = ERP_Price_pd_history.iloc[i,2]
            ERP_Price_pd_history.iloc[i, :] = target.iloc[0,:]
            ERP_Price_pd_new.drop(index=target.index[0],axis=0,inplace=True)

    if not ERP_Price_pd_new.empty:
        pd_ERP_Price = pd.concat([ERP_Price_pd_history,ERP_Price_pd_new])
    else:
        pd_ERP_Price = ERP_Price_pd_history

    ERP_Price_pd_history.to_excel(Results_FolderNameStr + FileNameStr_result_RefPrice_byERPnum)







print('OK')