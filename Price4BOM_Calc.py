import pandas as pd
import os

def ERP_BOM_Calc(ERP_num,BOM_FileFolder,Destination_FileFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename):
    BOM_FileName = BOM_FileFolder + ERP_num + '.xls'
    BOM_Price_FileName = Destination_FileFolder +'Price/'+ ERP_num +'_price.xls'
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
    BOM_df.to_excel(BOM_Price_FileName)


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

def Calc_Folder(BOM_FileFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename):
    if os.path.exists(BOM_FileFolder):
        for root, dirs, files in os.walk(BOM_FileFolder):
            Destination_FileFolder = BOM_FileFolder
            if not os.path.exists(BOM_FileFolder +'Price'):
                os.makedirs(BOM_FileFolder+'Price')
            for file in files :
                if os.path.splitext(file)[1].lower() == '.xls' :
                    #print(os.path.join(root,file))
                    BOM_subFileFolder=root+'/'
                    ERP_num = os.path.splitext(file)[0]
                    if not '_price' in ERP_num:
                        print('Calculating BOM Price for:',ERP_num,'....')
                        ERP_BOM_Calc(ERP_num, BOM_subFileFolder, Destination_FileFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename)


if False:
    ERP_num = 'P0402000175'
    BOM_FileFolder = './BOM/4D-888E-W/4T-13-24CH-BOM20190323-P0402000175/L1/'
    ERP_BOM_Calc(ERP_num,BOM_FileFolder)


PriceInfor_Filename = './Purchase_Rawdata/RefPrice_byERPnum_202011.xlsx'
PriceInfor_L2L3_Filename ='PriceInfor_L3L2_V3_0.xlsx'

BOM_FileFolder = './BOM/8D-989E/'
#subFolder = 'L4/'
#Calc_Folder(BOM_FileFolder+subFolder)
subFolder = 'L3/'
Calc_Folder(BOM_FileFolder+subFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename)
subFolder = 'L2/'
Calc_Folder(BOM_FileFolder+subFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename)
subFolder = 'L1/'
Calc_Folder(BOM_FileFolder+subFolder,PriceInfor_Filename,PriceInfor_L2L3_Filename)
