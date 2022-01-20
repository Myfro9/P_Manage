# P_Manage
### This Project is used for calculation of the product BOM price, analyze purchase information, from previous purchase database

### P_manage 程序运行说明：
***
1. 运行PurchaseinforV2.0.py，设定采购信息记录excel表，该程序可以整理出：
  - Result_byERPnum ： 根据ERP编码排列出所有的产品采购信息
  - Result_bySupplier ： 根据供应商名称列出所有的产品采购信息
  - SupplierSumPrice ： 根据供应商列表整理出相对应的每一个供应商所有历史采购金额
  - Result_RefPrice_byERPnum：根据ERP编码，列出所有产品的参考价格，这个参考价格是根据历史记录的所有该产品的采购价格，通过统计方式计算得到。该统计方法为：
      - 基于每一个ERP编码所对应的所有采购合同和所有的采购数量，求平均单价格和最低单价格
      - 针对每一份合同，计算其单价与平均单价格之间的差额
      - 根据差额对所有合同进行归类（数量、单价、数量占比例）
      - 剔除采购量小于10%且价格偏离平均值10%的那些差额项，对剔除后的结果再求均值，这个值作为最终的参考价格。
      - 如果出现高于平均价格50%，且采购量大于10%的情况，会报警1
      - 如果出现低于平均价格50%，且采购量大于10%的情况，会报警2

    通过采购得到历年所有的采购合同统计信息（内包含订单编号、ERP编码、厂家信息、单价、数量信息等），将该采购合同统计信息的文件名与之前的统计信息一起，输入到程序中。需要注意的是：多个采购合同统计信息文件，在签订合同的时间段上不要有重复，否则会造成信息重复统计，这个通过对多个文件进行简单的检查就可以避免。

    通过控制 if False 和 if True语句来分别执行不同的四个功能，生成上面提到的四个文件。
***
2. Price4BOM_Calc.py
  - ERP_BOM_Calc 函数：对某一个BOM文件进行分析，根据指定的参考价格信息，输出一个名为xxx_Price的价格信息文件，该函数被Calc_Folder函数调用。
  - Calc_Folder 函数： 对某个文件夹下所有指定文件（不包含 _Price）的BOM文件，自动批量计算价格信息。
  
  **使用方法：**
  
  **在文件中Line87修改采购统计信息表的文件PriceInfor_FileName的名称**
  
  **在文件中Line90修改BOM文件存放的文件夹BOM_FileFolder**
  
  **可以自动分析BOM的层数，最大支持L5**
  
  **运行程序**
  
  ---
  
  ![image](https://github.com/Myfro9/P_Manage/blob/Branch1/IMG/chart1.png)
  ---
  
# 主要功能
## 更新BOM
  
  1. 在 ./BOM_SVN/BOM 目录下从SVN上更新最新的BOM文件
  2. 直接运行 **BOM_update.py**，读取./BOM_SVN/BOM/下的BOM文件，自动生成L1～L5的等级的BOM
  
## 生成带价格的BOM
  运行 **Price4BOM_Calc.py** ， 计算./BOM.nosync/下的所有BOM的参考成本。运行前需要指定所有物料的历史采购参考价格
  
    Version = 1  # 0: only history record ; 1: 包含了近期的采购价格
    if Version ==1:
        PriceInfor_Filename = './Results/RefPrice_byERPnum.xlsx'  # 包含历史以及近期的采购信息
    else:
        PriceInfor_Filename = './Results/RefPrice_byERPnum_202104.xlsx'  # 仅历史采购信息
    PriceInfor_L2L3_Filename ='./Results/PriceInfor_L3L2_V3_0.xlsx'
  
## 拉缺料表
  1. 运行 **Purchase_Production_ListGen.py**，生成缺料表BOM清单Purchase_BOM，运行前需要指定：
  
    - BOM的位置
    BOM_folderNameStr = './BOM.nosync/'
    
    - 在调用的 **rd_DataBase.py** 中指定：生产计划单列表，采购信息，库存信息
    导入在产生产任务单(00):
    FileNameStr_PrjList = './Purchase_Rawdata/00在产任务单/近期在产生产任务单-2022-good.xls'
    
    导入库存信息(11)
    StockInfor_Filename = './Purchase_Rawdata/11库存记录/ERP现存量20211008.XLS'
    导入材料入库记录(12)
    FolderNameStr = './Purchase_Rawdata/12材料入库记录/'
    FileNameStr_instock = '采购入库单列表20211008.xls'
    导入材料出库记录(13)
    FolderNameStr = './Purchase_Rawdata/13材料出库记录/'
    FileNameStr_outstock = '材料出库单列表20211008.xls'
    导入委外加工出库记录(15)
    FolderNameStr = './Purchase_Rawdata/15委外材料出库记录/'
    FileNameStr_outsourcing = '委外材料出库单列表20211008.xls'
    导入委外加工入库记录(16)
    FolderNameStr = './Purchase_Rawdata/16委外产成品入库记录/'
    FileNameStr_outsourcingBack = '委外产成品入库单20211008.xls'
    
    导入所有的采购合同(21)
    FolderNameStr = './Purchase_Rawdata/21采购合同记录/'
    FileNameStr_contract = 'ERP20150101-20210420.xls'
    
    Contract_FolderNameStr = './Purchase_Rawdata/23询价结果/'
    Contract_FileNameStr1 = '合并汇总表-版本24-20211230.xlsx'
    
  2. 运行 **Quote_gen.py** ，生成询价表，在运行前需要指定：
    
    - 缺料表BOM清单，Purchase_BOM
    FileNameStr = 'PurchaseTable2022-01-12-20_17.xls'
    FolderNameStr = './Results/PurchaseBOM_2022-01-12-20_17/'

    - 输入需要生成BOM清单的生产计划单号
    Target = 'RW202201-A-1'
  
## 自动生成采购合同及付款申请
  运行 **Contract_gen.py** ,自动生成采购合同和付款申请合同，运行前需要指定：
  
    - 在'./PurchaseBOM_询价表_log.xlsx' 中，手动填写生产计划单对应的PurchaseBOM以及询价表的目录信息
    
    - 询价结果
    FolderNameStr = './Purchase_Rawdata/23询价结果/'
    Quote_result = '合并汇总表-版本27-20220110'

    - 供应商信息
    FolderNameStr = './Purchase_Rawdata/22供应商档案/'
    FileNameStr = '供应商档案 30-20220106.xlsx'

    - 工艺要求
    FolderNameStr0 = './Purchase_Rawdata/30工艺要求/'
    FileNameStr0 = '焊接工艺要求20211105.xls'

    - 需要生成采购合同的生产计划单号列表
    Tasklist = [...]
    Verlist = [...]
    
## 采购到货进度跟踪 
运行 **Schedule_gen.py** ，根据缺料表BOM清单和采购询价结果、采购到货信息等跟踪每一颗物料的到货情况
如果某个半成品的物料已经到齐，会在"生产任务派单表xxx.xlxs"中提示该半成品的物料都已经到齐
在运行前需要指定：

    -指定采购询价结果，（其中包含了合同、付款信息）
    FolderNameStr_quote = './Purchase_Rawdata/23询价结果/'
    FileNameStr_quote = '合并汇总表-版本19-20211222.xlsx'
    
    -在产生产任务单
    FileNameStr_PrjList = './Purchase_Rawdata/00在产任务单/近期在产生产任务单-8-11-good.xls'

    -导入采购到货记录
    FolderNameStr = './Purchase_Rawdata/24采购到货记录/'
    FileNameStr = '2021ERP到货单-20211210.xls'
    
    -导入供应商信息
    FolderNameStr = './Purchase_Rawdata/22供应商档案/'
    FileNameStr = '供应商档案 24-20211115.xlsx'
    
    -指定生产计划单号
    pd_Task = pd.DataFrame({'Task':[...] })