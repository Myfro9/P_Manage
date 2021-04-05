# P_Manage
This Project is used for calculation of the product BOM price, analyze purchase information, from previous purchase database

P_manage 程序运行说明：
1）运行PurchaseinforV2.0.py，设定采购信息记录excel表，该程序可以整理出：
  a)Result_byERPnum ： 根据ERP编码排列出所有的产品采购信息
  b)Result_bySupplier ： 根据供应商名称列出所有的产品采购信息
  c)SupplierSumPrice ： 根据供应商列表整理出相对应的每一个供应商所有历史采购金额
  d)Result_RefPrice_byERPnum：根据ERP编码，列出所有产品的参考价格，这个参考价格是根据历史记录的所有该产品的采购价格，通过统计方式计算得到。该统计方法为：
    i.基于每一个ERP编码所对应的所有采购合同和所有的采购数量，求平均单价格和最低单价格
    ii.针对每一份合同，计算其单价与平均单价格之间的差额
    iii.根据差额对所有合同进行归类（数量、单价、数量占比例）
    iv.剔除采购量小于10%且价格偏离平均值10%的那些差额项，对剔除后的结果再求均值，这个值作为最终的参考价格。
    v.如果出现高于平均价格50%，且采购量大于10%的情况，会报警1
    vi.如果出现低于平均价格50%，且采购量大于10%的情况，会报警2

  通过采购得到历年所有的采购合同统计信息（内包含订单编号、ERP编码、厂家信息、单价、数量信息等），将该采购合同统计信息的文件名与之前的统计信息一起，输入到程序中。需要注意的是：多个采购合同统计信息文件，在签订合同的时间段上不要有重复，否则会造成信息重复统计，这个通过对多个文件进行简单的检查就可以避免。

通过控制 if False 和 if True语句来分别执行不同的四个功能，生成上面提到的四个文件。

2）Price4BOM_Calc.py
    1)ERP_BOM_Calc 函数：对某一个BOM文件进行分析，根据指定的参考价格信息，输出一个名为xxx_Price的价格信息文件，该函数被Calc_Folder函数调用。
       i.需要在这个函数中修改采购统计信息表的文件PriceInfor_FileName的名称。
        ii.
    2)Calc_Folder 函数： 对某个文件夹下所有指定文件（不包含 _Price）的BOM文件，自动批量计算价格信息。
