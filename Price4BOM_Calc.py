import pandas as pd
import os
import time
import shutil

def ERP_BOM_Calc(ERP_num,BOM_FileFolder,Destination_FileFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename):
    BOM_FileName = BOM_FileFolder + ERP_num + '.xls'
    BOM_Price_FileName = Destination_FileFolder + ERP_num +'_price.xls'
    #PriceInfor_Filename = 'Purchase_0813A.xlsx'

    ## Read BOM xls file
    xls_df = pd.DataFrame(pd.read_excel(BOM_FileName))
    Name_BOM = xls_df.iloc[0,1]
    BOM_df = xls_df.iloc[:,[9,11,14]]
    #BOM_df = xls_df.loc[:, ['子件编码(cpscode)', '子件名称(cinvname)', '基本用量分子(ipsquantity)']]
    #BOM_df = xls_df.loc[:, ['子件编码', '子件名称', '基本用量分子']]
    BOM_df.columns= ['ERP编码','型号','数量']
    BOM_df['单价1']= None
    BOM_df['Name']= None

    #print(PriceInfor_L2L3_df)

    #print(ItemSelected_df)

    ## Read Price xls file
    xls2_df = pd.DataFrame(pd.read_excel(PriceInfor_Filename))
    PurchasePriceInfor_df = xls2_df.loc[:,['ERP','Ref_UnitPrice','Name']]
    PriceInfor_L2L3_df = pd.DataFrame(pd.read_excel(PriceInfor_L2L3_Filename)).iloc[:,[1,2,3]]
    PriceInfor_df = pd.concat([PurchasePriceInfor_df,PriceInfor_L2L3_df])
    for i in range(0,BOM_df.shape[0]):
        target_sr = BOM_df.iloc[i, :]
        if target_sr[0] in PriceInfor_df['ERP'].values:
            result_df = PriceInfor_df[PriceInfor_df['ERP']==target_sr[0]]  ## Search and Match
            BOM_df.iloc[i,3] = result_df.iloc[0,1]  ## Select useful item from search result
            BOM_df.iloc[i, 4] = result_df.iloc[0, 2]
        else:
            print('There is no ERP information about:', target_sr[0], 'in the Purchase information list')
            BOM_df.iloc[i, 3] = 0


    BOM_df['总价'] = BOM_df.apply(lambda x:x['数量'] * x['单价1'],axis=1) ## Calc between columns

    Sum = BOM_df['总价'].sum()
    BOM_df.loc['总计'] = None
    BOM_df.iloc[BOM_df.shape[0]-1,5] = Sum
    BOM_df_sorted = BOM_df.sort_values(by='总价', ascending=False)
    BOM_df_sorted.to_excel(BOM_Price_FileName)


    if ERP_num in PriceInfor_L2L3_df.iloc[:,0].values:
        target_index = PriceInfor_L2L3_df[PriceInfor_L2L3_df.iloc[:,0]==ERP_num].index.tolist()
        Recorded = PriceInfor_L2L3_df.iloc[target_index,1].values
        if round(Recorded[0],2) != round(Sum,2) :
            print('Warning: ', ERP_num, 'Price is different to recorded in L2L3 Price Information:', PriceInfor_L2L3_df.iloc[target_index,1].values)
            print('The current Price for: ', ERP_num, 'is:', Sum)
            PriceInfor_L2L3_df.iloc[target_index,1]= Sum
        else:
            print(ERP_num,'Price is the same as in L2L3 Price Information')
    else:
        PriceInfor_L2L3_df.loc[PriceInfor_L2L3_df.shape[0]] = [ERP_num,Sum,Name_BOM]
        # print(PriceInfor_L2L3_df)
        print(ERP_num,'Price is created in L2L3 Price Information')


    PriceInfor_L2L3_df.to_excel(PriceInfor_L2L3_Filename)

    ## Attach the L2L3 infor:




def Calc_Folder(BOM_FileFolder, Result_FileFolder, subFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename):
    if os.path.exists(BOM_FileFolder+'/'):
        if not os.path.exists(Result_FileFolder+'_Price/'):
            os.makedirs(Result_FileFolder+'_Price/')
        if not os.path.exists(Result_FileFolder + '_Price/'+subFolder):
            os.makedirs(Result_FileFolder + '_Price/'+subFolder)
        Destination_FileFolder = Result_FileFolder+'_Price/'+subFolder
        for root, dirs, files in os.walk(BOM_FileFolder+'/'+subFolder):
            for file in files:
                if os.path.splitext(file)[1].lower() == '.xls':
                    BOM_subFileFolder = root
                    ERP_num = os.path.splitext(file)[0]
                    if not '_price' in ERP_num:
                        ERP_BOM_Calc(ERP_num, BOM_subFileFolder, Destination_FileFolder, PriceInfor_Filename,PriceInfor_L2L3_Filename)



Task = 'P0402000368'   # ERP编码：可以单独输入ERP编码； 'all'： 所有的BOM一起计算
Version = 1  # 0: only history record ; 1: 包含了近期的采购价格

if Version ==1:
    PriceInfor_Filename = './Results/RefPrice_byERPnum.xlsx'
else:
    PriceInfor_Filename = './Results/RefPrice_byERPnum_202104.xlsx'
PriceInfor_L2L3_Filename ='./Results/PriceInfor_L3L2_V3_0.xlsx'
BOM_FileFolder = './BOM/'
Result_FileFolder0 = './Results/BOM_Price'
TimeStr = time.strftime("%Y-%m-%d-%H_%M", time.localtime(time.time()))
Result_FileFolder = Result_FileFolder0 + '/BOM_Price_' + TimeStr

assert not os.path.exists(Result_FileFolder), '目标文件夹已经存在了,再等几分钟！'
os.makedirs(Result_FileFolder)
ERP = ''
initial_status = 1
for root, dirs, files in os.walk(BOM_FileFolder):
    for file in files:
        print(root + '/' + file)
        if file.split('.')[-1].lower() == 'xls' or file.split('.')[-1].lower() == 'xlsx':
            ERP_tmp = root.split('/')[2]
            if ERP != ERP_tmp:
                if initial_status == 0:
                    L1_L2_L3_L4 = L0
                    if ERP == Task or Task =='all':
                        for i in range(L1_L2_L3_L4, 0, -1):
                            subFolder = 'L' + str(i) + '/'
                            Calc_Folder(BOM_FileFolder + ERP, Result_FileFolder + '/'+ ERP, subFolder, PriceInfor_Filename,
                                        PriceInfor_L2L3_Filename)
                    #print(ERP)
                    #print(L1_L2_L3_L4)
                initial_status = 0
                ERP = ERP_tmp
                L0 = 1
                print(ERP_tmp)
            if root.split('/')[3] == 'L2':
                if L0 <=2:
                    L0 = 2
            elif root.split('/')[3] == 'L3':
                if L0 <=3:
                    L0 = 3
            elif root.split('/')[3] == 'L4':
                    L0 = 4
            print(L0)



'''L1_L2_L3_L4 = 4  # Specify the maximum layers in the BOM structure.

for i in range(L1_L2_L3_L4,0,-1):
    subFolder = 'L'+str(i)+'/'
    Calc_Folder(BOM_FileFolder, Result_FileFolder, subFolder, PriceInfor_Filename, PriceInfor_L2L3_Filename)'''


