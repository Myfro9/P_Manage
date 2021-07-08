import pandas as pd
import os
import datetime
import time
import xlwings as xw

####### 导入询价结果
FolderNameStr = './Purchase_Rawdata/23询价结果/'
FileNameStr = '汇总YF06AB、GC06AB、RW06DEF12G12H1-2-20210707.xlsx'
pd_Quote_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr,sheet_name='操作'))

FolderNameStr_result = './Results/汇总YF06AB、GC06AB、RW06DEF12G12H1-2-20210707/'
if not os.path.exists(FolderNameStr_result):
    os.makedirs(FolderNameStr_result)


####### 导入供应商信息
FolderNameStr = './Purchase_Rawdata/22供应商档案/'
FileNameStr = '供应商档案 1-20210707.xlsx'
pd_Supplier_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr))

####### 打开合同模版
FolderNameStr = './Purchase_Rawdata/'
FileNameStr_Contract0 = 'Contract Template.xlsx'
FileNameStr_Contract1 = 'Contract Template_long.xlsx'
FileNameStr1 = 'Contract target.xlsx'

####### 打开付款模版
FileNameStr_Pyament = '付款单模板.xlsx'


app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False  # 是否实时刷新excel程序的显示内容


####### 合同存储


def int2Chnese(number, recursive_depth=0):
    str_number = str(number)
    if len(str_number) > 4:
        str_number = str_number[-4:]
    bits = "零 壹 贰 叁 肆 伍 陆 柒 捌 玖".split(" ")

    units = " 拾 佰 仟".split(" ")
    large_unit = ' 万 亿 万'.split(" ")  # 可扩展,以万为单位
    number_len = len(str_number)
    result = ""

    for i in range(number_len):
        result += bits[int(str_number[i])]
        if str_number[i] != "0":
            result += units[number_len - i - 1]

    # 去除连续的零
    while "零零" in result:
        result = result.replace("零零", "零")
    # 去除尾部的零
    if result[-1] == "零":
        result = result[:-1]
    # 调整10~20之间的数
    if result[:2] == "一十":
        result = result[1:]
    # 字符串连接上大单位
    result += large_unit[recursive_depth]

    # 判断是否递归
    if len(str(number)) > 4:
        recursive_depth += 1
        return int2Chnese(str(number)[:-4], recursive_depth) + result
    else:
        return result



def supplier_Inforsearch(kwd,pd_SupplierInfor):
    infor = pd_SupplierInfor[pd_SupplierInfor['供应商名称'].str.contains(kwd)]
    #print(infor.shape[0])
    if infor.shape[0] > 1:
        print(kwd + ' 有多条供应商记录')
        infor = infor.iloc[infor.shape[0]-1,:]
    return infor

#print(int2Chnese(3467))
TimeStr = time.strftime("%Y-%m-%d", time.localtime(time.time()))
pd_grouped_Quote = pd_Quote_Infor.groupby('生产计划单号')
for item, item_df in pd_grouped_Quote:
    pd_grouped2_Quote = item_df.groupby('渠道')
    task = item
    for item2, item2_df in pd_grouped2_Quote:
        supplier_kwd = item2.strip()
        supplier_infor = supplier_Inforsearch(supplier_kwd, pd_Supplier_Infor)
        # assert not supplier_infor.empty, '供应商档案没有{}'.format(supplier_kwd)
        if supplier_infor.empty:
            print('Warning: 供应商档案没有 {} 的信息'.format(supplier_kwd))
        else:
            if item2_df.shape[0] < 11:
                wb = app.books.open(FolderNameStr + FileNameStr_Contract0)
                ws_sheet0 = wb.sheets[0]
                ws_sheet1 = wb.sheets[1]
                offset = 0
            else:
                wb = app.books.open(FolderNameStr + FileNameStr_Contract1)
                ws_sheet0 = wb.sheets[0]
                ws_sheet1 = wb.sheets[1]
                offset = 15
            assert item2_df.shape[0]<26 , '采购列表太长，超过25条了！'
            supplier_infor = supplier_infor.iloc[0,:]
            payment_infor = supplier_infor['账期']
            for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                ERP = pd_Quote_Infor.loc[item2_df.index[i],'ERP编码']
                Type = pd_Quote_Infor.loc[item2_df.index[i],'型号']
                Class = pd_Quote_Infor.loc[item2_df.index[i],'供应商类别']
                num = pd_Quote_Infor.loc[item2_df.index[i],'报价数量']
                unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                if num != None and num > 0:
                    if i ==0:
                        infor0 = '卖方： ' + supplier_infor['供应商名称'] + '（以下简称乙方）'
                        infor1 = '卖方（乙方）： ' + supplier_infor['供应商名称']
                        if str(supplier_infor['地址']) != 'nan':
                            infor2 = '地址： ' + str(supplier_infor['地址'])
                        else:
                            infor2 = '地址： '

                        if str(supplier_infor['联系人']) != 'nan':
                            infor3 = '经办人： ' + str(supplier_infor['联系人'])
                        else:
                            infor3 = '经办人： '

                        if str(supplier_infor['电话']) != 'nan':
                            infor4 = '经办人联系电话/传真： ' + str(supplier_infor['电话'])
                        else:
                            infor4 = '经办人联系电话/传真： '
                        if  str(supplier_infor['传真']) != 'nan':
                            infor4 = infor4 + ' / ' + str(supplier_infor['传真'])

                        if str(supplier_infor['开户行']) !='nan' and str(supplier_infor['账号']) != 'nan':
                            infor5 = '开户行及账号： ' + str(supplier_infor['开户行']) + ' / ' + str(supplier_infor['账号'])
                        else:
                            infor5 = '开户行及账号： '
                        ws_sheet0.range('H6').value = '签订日期： '+ TimeStr
                        ws_sheet0.range('A7').value = infor0
                        ws_sheet0.range('E'+str(60+offset)).value = infor1
                        ws_sheet0.range('A8').value = infor2
                        ws_sheet0.range('E'+str(62+offset)).value = infor3
                        ws_sheet0.range('E'+str(63+offset)).value = infor4
                        ws_sheet0.range('E'+str(64+offset)).value = infor5
                        ws_sheet1.range('A2').value = task

                        ws_sheet0.range('B12').value = Type
                        ws_sheet0.range('D12').value = ERP
                        ws_sheet0.range('E12').value = num
                        ws_sheet0.range('F12').value = unit_price
                        ws_sheet0.range('G12').value = total_price
                        ws_sheet0.range('H12').value = '现货'

                    else:
                        idx = str(12+i)
                        ws_sheet0.range('B'+idx).value = Type
                        ws_sheet0.range('D'+idx).value = ERP
                        ws_sheet0.range('E'+idx).value = num
                        ws_sheet0.range('F'+idx).value = unit_price
                        ws_sheet0.range('G'+idx).value = total_price
                        ws_sheet0.range('H'+idx).value = '现货'

            sumed_price =ws_sheet0.range('G'+str(22+offset)).value
            sumed_price_Chinese = int2Chnese(int(sumed_price))
            ws_sheet0.range('A'+str(22+offset)).value = '以上单价含13%增值税，大写金额：'+ sumed_price_Chinese + '元整'
            # 帐期
            if payment_infor == '预付30%+票到30天':
                ws_sheet0.range('A'+str(42+offset)).value = '1、支付方式：30%预付款，货到票到30天月结'
            elif payment_infor == '款到发货':
                ws_sheet0.range('A'+str(42+offset)).value = '1、支付方式：款到发货'
            elif payment_infor == '票到30天':
                ws_sheet0.range('A'+str(42+offset)).value = '1、支付方式：货到票到30天月结'
            elif payment_infor == '货到付款':
                ws_sheet0.range('A'+str(42+offset)).value = '1、支付方式：货到付款'
            elif payment_infor == '预付30 %，付清尾款发货':
                ws_sheet0.range('A' + str(42 + offset)).value = '1、支付方式：预付30 %，付清尾款发货'
            elif payment_infor == '月结30天':
                ws_sheet0.range('A' + str(42 + offset)).value = '1、支付方式：货到月结30天'
            else:
                ws_sheet0.range('A'+str(42+offset)).value = '1、支付方式：货到月结'
            if str(Class) == 'nan' :
                FileNameStr_result =  '_' + task + '_' + item2 + '_合同.xlsx'
                FileNameStr_result_old =  '_' + task + '_' + item2 + '_合同_old.xlsx'
            else:
                FileNameStr_result = Class + '_' + task + '_' + item2 + '_合同.xlsx'
                FileNameStr_result1 = Class + '_' + task + '_' + item2 + '_合同_old.xlsx'
            if os.path.exists(FolderNameStr_result + FileNameStr_result):
                if os.path.exists(FolderNameStr_result + FileNameStr_result_old):
                    os.remove(FolderNameStr_result + FileNameStr_result_old)
                os.rename(FolderNameStr_result + FileNameStr_result,FolderNameStr_result + FileNameStr_result_old)


            wb.save(FolderNameStr_result + FileNameStr_result)
            wb.close()


            print(task + '_' + item2)



app.quit()
print('test1')