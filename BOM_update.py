import pandas as pd
import os
import datetime
import time
import shutil

###############################################################
### 根据SVN上的EXCEL文件，自动生成L1，L2... L6层次的结构化BOM文件 ###
###############################################################

TimeStr = time.strftime("%Y-%m-%d-%H_%M", time.localtime(time.time()))
BOMFolderNameStr = './BOM_SVN/BOM/'
BOMlist_FileNameStr = './BOM_SVN/BOM_ERP编码对照表.xlsx'

Dest_BOMFolderNameStr = './BOM_SVN/Auto_gen/' + TimeStr +'/'
pd_BOMlist = pd.DataFrame(columns=['ERP编码','名称','型号'])

assert not os.path.exists(Dest_BOMFolderNameStr), '目标文件夹已经存在了,再等几分钟！'
os.makedirs(Dest_BOMFolderNameStr)

def Build_Child_Subfolder(ERPnum,ser_ERPlist,BOMFolderNameStr,Dest_folder):
    pd_subERP = pd.DataFrame(columns=['ERP'])
    if ERPnum in ser_ERPlist.values:
        if not os.path.exists(Dest_folder):
            os.makedirs(Dest_folder)
        target_file = ERPnum + '.XLS'
        assert os.path.exists(BOMFolderNameStr+target_file), 'ERP编码：{}文件不存在！'.format(BOMFolderNameStr+target_file)
        shutil.copyfile(BOMFolderNameStr+target_file, Dest_folder+target_file)
        pd_subERP = pd.DataFrame(pd.read_excel(BOMFolderNameStr+target_file))

    return pd_subERP

if False:  # 拷贝历史BOM到新的BOM文件夹里
    for root,dirs, files in os.walk('./BOM/'):
        for file in files:
            if file.split('.')[-1].lower() == 'xls' or file.split('.')[-1].lower() == 'xlsx':
                print(root+file)
                print(BOMFolderNameStr + file)
                shutil.copyfile(root+'/'+file, BOMFolderNameStr + file)


#####  生成BOM文件和ERP编码的对应说明文件
for root, dirs, files in os.walk(BOMFolderNameStr):
    for file in files:
        #print(file)
        if file.split('.')[-1].lower() == 'xls' or file.split('.')[-1].lower() == 'xlsx':
            BOM_file = root  + file.lower()
            pd_BOMfile = pd.DataFrame(pd.read_excel(BOM_file))
            #ERP = pd_BOMfile.iloc[0,0]
            ERP = file.split('.')[0].upper()
            Name = pd_BOMfile.iloc[0,1]
            Type = pd_BOMfile.iloc[0,2]
            pd_BOMinfo = pd.DataFrame([[ERP,Name,Type]],columns=['ERP编码','名称','型号'])
            if pd_BOMlist.empty:
                pd_BOMlist = pd_BOMinfo
            else:
                pd_BOMlist = pd.concat([pd_BOMlist,pd_BOMinfo])
                inx = pd_BOMlist.shape[0]
                pd_BOMlist.index=range(inx)
pd_BOMlist.to_excel(BOMlist_FileNameStr)
ser_ERPlist = pd_BOMlist['ERP编码']


####  自动生成L1，L2... L6层次的结构化BOM文件
for root, dirs, files in os.walk(BOMFolderNameStr):
    for file in files:
        #print(file)
        if file.split('.')[-1].lower() == 'xls' or file.split('.')[-1].lower() == 'xlsx':
            BOM_file = root + file.lower()
            pd_BOMfile = pd.DataFrame(pd.read_excel(BOM_file))
            print(BOM_file)
            ERP = file.split('.')[0].upper()
            # 根据当前文件的ERP编码，建立L1目录
            print(ERP)
            if not os.path.exists(Dest_BOMFolderNameStr +'/' + ERP ):
                os.makedirs(Dest_BOMFolderNameStr +'/' + ERP)
                os.makedirs(Dest_BOMFolderNameStr +'/' + ERP +'/L1/')
                Dest_BOM_file = Dest_BOMFolderNameStr +'/' + ERP +'/L1/' + file.lower()
                shutil.copyfile(BOM_file, Dest_BOM_file)
            # 根据当前文件内部的每一个部件的ERP编码，建立L2目录
            for i in range(pd_BOMfile.shape[0]):
                ERPnumL1 = pd_BOMfile.loc[pd_BOMfile.index[i],'子件编码(cpscode)']
                pd_subERPL2 = Build_Child_Subfolder(ERPnumL1,ser_ERPlist,BOMFolderNameStr,Dest_BOMFolderNameStr +'/' + ERP +'/L2/')
                # 继续搜索L2目录下的ERP文件，看看是否存在L3
                if not pd_subERPL2.empty:
                    for j in range(pd_subERPL2.shape[0]):
                        ERPnumL2 = pd_subERPL2.loc[pd_subERPL2.index[j],'子件编码(cpscode)']
                        pd_subERPL3 = Build_Child_Subfolder(ERPnumL2, ser_ERPlist,BOMFolderNameStr,
                                                            Dest_BOMFolderNameStr + '/' + ERP + '/L3/')
                        #继续搜索L3目录下的ERP文件，看看是否存在L4
                        if not pd_subERPL3.empty:
                            for k in range(pd_subERPL3.shape[0]):
                                ERPnumL3 = pd_subERPL3.loc[pd_subERPL3.index[k],'子件编码(cpscode)']
                                pd_subERPL4 = Build_Child_Subfolder(ERPnumL3, ser_ERPlist,BOMFolderNameStr,
                                                                    Dest_BOMFolderNameStr + '/' + ERP + '/L4/')

                                # 继续搜索L4目录下的ERP文件，看看是否存在L5
                                if not pd_subERPL4.empty:
                                    for k4 in range(pd_subERPL4.shape[0]):
                                        ERPnumL4 = pd_subERPL4.loc[pd_subERPL4.index[k4],'子件编码(cpscode)']
                                        pd_subERPL5 = Build_Child_Subfolder(ERPnumL4, ser_ERPlist,BOMFolderNameStr,
                                                                            Dest_BOMFolderNameStr + '/' + ERP + '/L5/')

                                        # 继续搜索L5目录下的ERP文件，看看是否存在L6
                                        if not pd_subERPL5.empty:
                                            for k5 in range(pd_subERPL5.shape[0]):
                                                ERPnumL5 = pd_subERPL5.loc[pd_subERPL5.index[k5],'子件编码(cpscode)']
                                                pd_subERPL6 = Build_Child_Subfolder(ERPnumL5, ser_ERPlist,BOMFolderNameStr,
                                                                                    Dest_BOMFolderNameStr + '/' + ERP + '/L6/')


####  将生成好的结构化BOM，拷贝到目标文件夹中。
if os.path.exists('./BOM_SVN/'+ 'backup_' + TimeStr +'/'):
    os.makedirs('./BOM_SVN/'+ 'backup_' + TimeStr +'/')
shutil.copytree('./BOM/', './BOM_SVN/'+ 'backup_' + TimeStr +'/')
shutil.rmtree('./BOM/')
shutil.copytree(Dest_BOMFolderNameStr, './BOM/')
