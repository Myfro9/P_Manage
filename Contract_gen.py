import pandas as pd
import os
import datetime
import time
import xlwings as xw
import shutil

function = 1  # 0: 合同生成； 1：付款单生成
Target_Task = 'RW202106-E-1'    # 'all' 或者特定生产计划单号

####### 导入询价结果
FolderNameStr = './Purchase_Rawdata/23询价结果/'
FileNameStr = '汇总YF06AB、GC06AB、RW06DEF12G12H1-13-20210730.xlsx'
pd_Quote_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr,sheet_name='操作'))

FolderNameStr_result = './Results/汇总YF06AB、GC06AB、RW06DEF12G12H1-13-20210730/'
if not os.path.exists(FolderNameStr_result):
    os.makedirs(FolderNameStr_result)


####### 导入供应商信息
FolderNameStr = './Purchase_Rawdata/22供应商档案/'
FileNameStr = '供应商档案 9-20210727.xlsx'
pd_Supplier_Infor = pd.DataFrame(pd.read_excel(FolderNameStr+FileNameStr))

####### 打开合同模版
FolderNameStr = './Purchase_Rawdata/'
FileNameStr_Contract0 = 'Excel_templates/Contract Template.xlsx'
FileNameStr_Contract1 = 'Excel_templates/Contract Template_long.xlsx'
FileNameStr_Contract2 = 'Excel_templates/Contract Template_longest.xlsx'
FileNameStr1 = 'Excel_templates/Contract target.xlsx'

####### 打开付款模版
FileNameStr_Pyament = 'Excel_templates/付款单模板.xlsx'
FileNameStr_Pyament1 = 'Excel_templates/付款单模板_long.xlsx'
FileNameStr_Pyament2 = 'Excel_templates/付款单模板_longest.xlsx'

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False  # 是否实时刷新excel程序的显示内容


####### 合同汇总列表
FileNameStr_ContractList = 'Excel_templates/合同申请单模版.xlsx'
FileNameStr_PyamentList = 'Excel_templates/预付款申请单模版.xlsx'


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
    #print('Test:',kwd)
    infor = pd_SupplierInfor[pd_SupplierInfor['供应商名称'].str.contains(kwd).fillna(False)]
    #print(infor.shape[0])
    if infor.shape[0] > 1:
        print(kwd + ' 有' + str(infor.shape[0]) + ' 条供应商记录')
        infor = infor.iloc[[-1],:]
        #infor1 = infor.iloc[-1,:]
        #infor = pd.DataFrame(infor1.values,index=infor1.index).T
        #print(infor)
    return infor

def check_supplier_vld(pd_item2):
    pd_match =pd_item2[pd_item2['报价数量'] >0]
    if pd_match.empty:
        vld = 0
    else:
        vld =1
    return vld

def my_wbsave(wb_wc, FileNameStr):
    #wb.save(FolderNameStr_result + task + '付款单/' + FileNameStr_result_payment)
    wb_wc.save('./results/temp_result.xlsx')
    shutil.copyfile('./results/temp_result.xlsx',FileNameStr)

#print(int2Chnese(3467))
TimeStr = time.strftime("%Y-%m-%d", time.localtime(time.time()))
pd_grouped_Quote = pd_Quote_Infor.groupby('生产计划单号')

for item, item_df in pd_grouped_Quote:
    pd_grouped2_Quote = item_df.groupby('渠道')
    task = item
    if Target_Task == 'all' or Target_Task == task:
        if function ==0: #合同生成
            wc = app.books.open(FolderNameStr + FileNameStr_ContractList)   # wc是用于某个生产计划单的合同或付款单信息汇总，wb是用于每一份合同或付款单
        else: # 付款单生成
            wc = app.books.open(FolderNameStr + FileNameStr_PyamentList)
        wc_sheet0 = wc.sheets[0]
        wc_sheet0.range('B43').value = task
        wc_sheet0.range('C3').value = '提交日期： ' + TimeStr

        j =0   # j 是用来计数不同的供应商合同
        k =0   # k 是用来技术不同供应商的付款单
        for item2, item2_df in pd_grouped2_Quote:
            supplier_kwd = item2.strip()
            supplier_infor = supplier_Inforsearch(supplier_kwd, pd_Supplier_Infor)
            supplier_vld = check_supplier_vld(item2_df)
            # assert not supplier_infor.empty, '供应商档案没有{}'.format(supplier_kwd)
            if supplier_infor.empty:
                print('Warning: 供应商档案没有 {} 的信息，跳过'.format(supplier_kwd))
            elif supplier_vld == 0:
                print('Warning: 供应商{}采购量为空,跳过'.format(supplier_kwd))
            else:
                if function == 0: # 合同生成
                    if item2_df.shape[0] < 11:
                        wb = app.books.open(FolderNameStr + FileNameStr_Contract0)
                        ws_sheet0 = wb.sheets[0]
                        ws_sheet1 = wb.sheets[1]
                        offset = 0
                    elif item2_df.shape[0] < 26:
                        wb = app.books.open(FolderNameStr + FileNameStr_Contract1)
                        ws_sheet0 = wb.sheets[0]
                        ws_sheet1 = wb.sheets[1]
                        offset = 15
                    else:
                        wb = app.books.open(FolderNameStr + FileNameStr_Contract2)
                        ws_sheet0 = wb.sheets[0]
                        ws_sheet1 = wb.sheets[1]
                        offset = 30
                    assert item2_df.shape[0]<41 , '采购列表太长，超过40条了！'
                else: # 付款单生成
                    if item2_df.shape[0] < 11:
                        wb = app.books.open(FolderNameStr + FileNameStr_Pyament)
                        ws_sheet0 = wb.sheets[0]
                        offset = 0
                    elif item2_df.shape[0] < 26:
                        wb = app.books.open(FolderNameStr + FileNameStr_Pyament1)
                        ws_sheet0 = wb.sheets[0]
                        offset = 15
                    else:
                        wb = app.books.open(FolderNameStr + FileNameStr_Pyament2)
                        ws_sheet0 = wb.sheets[0]
                        offset = 30
                    assert item2_df.shape[0]<41 , '付款列表太长，超过40条了！'


                supplier_infor = supplier_infor.iloc[0,:]
                payment_infor = supplier_infor['账期']
                for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                    ERP = pd_Quote_Infor.loc[item2_df.index[i],'ERP编码']
                    Type = pd_Quote_Infor.loc[item2_df.index[i],'型号']
                    Class = pd_Quote_Infor.loc[item2_df.index[i],'供应商类别']
                    num = pd_Quote_Infor.loc[item2_df.index[i],'报价数量']
                    unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                    total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                    if num != None and str(num) != 'nan' and num > 0:
                        if i ==0:
                            if function == 0:  # 合同生成
                                infor0 =  supplier_infor['供应商名称'] + '（以下简称乙方）'
                                infor1 = '卖方（乙方）： ' + supplier_infor['供应商名称']
                                if str(supplier_infor['地址']) != 'nan':
                                    infor2 = str(supplier_infor['地址'])
                                else:
                                    infor2 = ''

                                if str(supplier_infor['联系人']) != 'nan':
                                    infor3 = '经办人： ' + str(supplier_infor['联系人'])
                                else:
                                    infor3 = '经办人： '

                                if str(supplier_infor['电话']) != 'nan':
                                    infor4 = '经办人联系电话/传真： ' + str(supplier_infor['电话'])
                                else:
                                    infor4 = '经办人联系电话/传真： '
                                if str(supplier_infor['传真']) != 'nan':
                                    infor4 = infor4 + ' / ' + str(supplier_infor['传真'])

                                if str(supplier_infor['开户行']) !='nan' and str(supplier_infor['账号']) != 'nan':
                                    #infor5 = '开户行及账号： ' + str(supplier_infor['开户行']) + ' / ' + str(supplier_infor['账号'])
                                    infor5 = '开户行及账号： '
                                else:
                                    infor5 = '开户行及账号： '


                                ws_sheet0.range('H6').value =  TimeStr
                                ws_sheet0.range('B7').value = infor0
                                ws_sheet0.range('E'+str(60+offset)).value = infor1
                                ws_sheet0.range('B8').value = infor2
                                ws_sheet0.range('E'+str(62+offset)).value = infor3
                                ws_sheet0.range('E'+str(63+offset)).value = infor4
                                ws_sheet0.range('E'+str(64+offset)).value = infor5
                                ws_sheet0.range('H5').value = task

                                ws_sheet0.range('B12').value = Type
                                ws_sheet0.range('D12').value = ERP
                                ws_sheet0.range('E12').value = num
                                ws_sheet0.range('F12').value = unit_price
                                ws_sheet0.range('G12').value = total_price
                                ws_sheet0.range('H12').value = ' '
                            else: # 付款单生成
                                infor0 =  supplier_infor['供应商名称']
                                if str(supplier_infor['开户行']) != 'nan' :
                                    infor5 = str(supplier_infor['开户行'])
                                else:
                                    infor5 = '需补充对方开户行信息'

                                if str(supplier_infor['账号']) != 'nan' :
                                    infor6 = str(supplier_infor['账号'])
                                else:
                                    infor6 = ''

                                ws_sheet0.range('A' + str(30 + offset)).value = infor0
                                ws_sheet0.range('A' + str(31 + offset)).value = infor5
                                ws_sheet0.range('A' + str(32 + offset)).value = infor6
                                ws_sheet0.range('D' + str(29 + offset)).value = TimeStr
                                ws_sheet0.range('C' + str(35 + offset)).value = task
                                if 'RW' in task:
                                    ws_sheet0.range('C' + str(34 + offset)).value = '计划部生产计划单'
                                elif 'YF' in task:
                                    ws_sheet0.range('C' + str(34 + offset)).value = '计划部研发计划单'
                                elif 'GC' in task:
                                    ws_sheet0.range('C' + str(34 + offset)).value = '工程部工程计划单'
                                else:
                                    ws_sheet0.range('C' + str(34 + offset)).value = ''
                                ws_sheet0.range('A8').value = Type
                                ws_sheet0.range('B8').value = num
                                ws_sheet0.range('C8').value = unit_price
                                ws_sheet0.range('E8').value = total_price
                        else:
                            if function == 0:  # 合同生成
                                idx = 12+i
                                ws_sheet0.range('B'+str(idx)).value = Type
                                ws_sheet0.range('D'+str(idx)).value = ERP
                                ws_sheet0.range('E'+str(idx)).value = num
                                ws_sheet0.range('F'+str(idx)).value = unit_price
                                ws_sheet0.range('G'+str(idx)).value = total_price
                                #ws_sheet0.range('H'+str(idx)).value = ws_sheet0.range('H'+str(idx-1)).value
                                # 交期
                                if str(Class) == 'mechanic':
                                    deliverTime = 21
                                else:
                                    deliverTime = 14
                                # 帐期
                                if payment_infor == '预付30%+票到30天':
                                    ws_sheet0.range('H'+str(idx)).value = '预付款后' + str(deliverTime) + '天'
                                elif payment_infor == '款到发货':
                                    ws_sheet0.range('H'+str(idx)).value = '款到发货'
                                elif payment_infor == '票到30天':
                                    ws_sheet0.range('H'+str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                elif payment_infor == '货到付款':
                                    ws_sheet0.range('H'+str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                elif payment_infor == '预付30 %，付清尾款发货':
                                    ws_sheet0.range('H'+str(idx)).value = '预付款后' + str(deliverTime) + '天'
                                elif payment_infor == '月结30天':
                                    ws_sheet0.range('H'+str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                else:
                                    ws_sheet0.range('H'+str(idx)).value = '合同签订后' + str(deliverTime) + '天'


                            else: # 付款单生成
                                idx = 8+i
                                ws_sheet0.range('A'+str(idx)).value = Type
                                ws_sheet0.range('B'+str(idx)).value = num
                                ws_sheet0.range('C'+str(idx)).value = unit_price
                                ws_sheet0.range('E'+str(idx)).value = total_price

                if function ==0:
                    sumed_price =ws_sheet0.range('G'+str(22+offset)).value
                    sumed_price_Chinese = int2Chnese(int(sumed_price))
                    ws_sheet0.range('A'+str(22+offset)).value = '以上单价含13%增值税，大写金额：'+ sumed_price_Chinese + '元整'
                    # 交期
                    if str(Class) == 'mechanic':
                        deliverTime = 21
                    else:
                        deliverTime =14
                    # 帐期
                    if payment_infor == '预付30%+票到30天':
                        ws_sheet0.range('B'+str(42+offset)).value = '30%预付款，货到票到30天月结'
                        ws_sheet0.range('H12').value = '预付款后'+ str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '预付款后'+ str(deliverTime) + '天发货'
                    elif payment_infor == '款到发货':
                        ws_sheet0.range('B'+str(42+offset)).value = '款到发货'
                        ws_sheet0.range('H12').value = '款到发货'
                        ws_sheet0.range('B' + str(36 + offset)).value = '款到发货'
                    elif payment_infor == '票到30天':
                        ws_sheet0.range('B'+str(42+offset)).value = '货到票到30天月结'
                        ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '合同签订后' + str(deliverTime) + '天发货'
                    elif payment_infor == '货到付款':
                        ws_sheet0.range('B'+str(42+offset)).value = '货到付款'
                        ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '合同签订后' + str(deliverTime) + '天发货'
                    elif payment_infor == '预付30 %，付清尾款发货':
                        ws_sheet0.range('B' + str(42 + offset)).value = '预付30 %，付清尾款发货'
                        ws_sheet0.range('H12').value = '预付款后' + str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '预付款后' + str(deliverTime) + '天发货'
                    elif payment_infor == '月结30天':
                        ws_sheet0.range('B' + str(42 + offset)).value = '货到月结30天'
                        ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '合同签订后' + str(deliverTime) + '天发货'
                    else:
                        ws_sheet0.range('B'+str(42+offset)).value = '货到月结'
                        ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                        ws_sheet0.range('B' + str(36 + offset)).value = '合同签订后' + str(deliverTime) + '天发货'




                    if str(Class) == 'nan' :
                        FileNameStr_result =  '_' + task + '_' + item2 + '_合同.xlsx'
                        FileNameStr_result_old =  '_' + task + '_' + item2 + '_合同_old.xlsx'
                    else:
                        FileNameStr_result = Class + '_' + task + '_' + item2 + '_合同.xlsx'
                        FileNameStr_result_old = Class + '_' + task + '_' + item2 + '_合同_old.xlsx'

                    if not os.path.exists(FolderNameStr_result + task + '合同/'):
                        os.makedirs(FolderNameStr_result + task + '合同/')
                    if os.path.exists(FolderNameStr_result + task + '合同/' +FileNameStr_result):
                        if os.path.exists(FolderNameStr_result + task + '合同/'+ FileNameStr_result_old):
                            os.remove(FolderNameStr_result + task + '合同/'+ FileNameStr_result_old)
                        os.rename(FolderNameStr_result + task + '合同/'+ FileNameStr_result,FolderNameStr_result + task + '合同/'+ FileNameStr_result_old)
                    #wb.save(FolderNameStr_result + task + '合同/'+ FileNameStr_result)
                    my_wbsave(wb, FolderNameStr_result + task + '合同/'+ FileNameStr_result)
                    wc_sheet0.range('E'+str(6+j)).value = ws_sheet0.range('G'+str(22+offset)).value  #合同金额
                    wc_sheet0.range('B' + str(6 + j)).value = supplier_infor['供应商名称']            # 供应商名称
                    wc_sheet0.range('F' + str(6 + j)).value = ws_sheet0.range('H12').value  # 付款方式
                    if item2_df.shape[0] >1:
                        wc_sheet0.range('C' + str(6 + j)).value = str(ws_sheet0.range('B12').value) + '  等'+ str(item2_df.shape[0]) + '种产品'
                    else:
                        wc_sheet0.range('C' + str(6 + j)).value = ws_sheet0.range('B12').value
                    wc_sheet0.range('D' + str(6 + j)).value = item2_df.shape[0]
                    j = j + 1
                    print(task + '_' + item2)

                else: #付款单生成
                    if str(Class) == 'nan' :
                        FileNameStr_result_payment =  '_' + task + '_' + item2 + '_付款.xlsx'
                        FileNameStr_result_payment_old =  '_' + task + '_' + item2 + '_付款_old.xlsx'
                    else:
                        FileNameStr_result_payment = Class + '_' + task + '_' + item2 + '_付款.xlsx'
                        FileNameStr_result_payment_old = Class + '_' + task + '_' + item2 + '_付款_old.xlsx'

                    if not os.path.exists(FolderNameStr_result + task + '付款单/'):
                        os.makedirs(FolderNameStr_result + task + '付款单/')
                    if os.path.exists(FolderNameStr_result + task + '付款单/' +FileNameStr_result_payment):
                        if os.path.exists(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment_old):
                            os.remove(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment_old)
                        os.rename(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment,FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment_old)

                    needPay =0
                    if payment_infor == '预付30%+票到30天':
                        ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                        ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                        #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                        my_wbsave(wb, FolderNameStr_result + task + '付款单/' + FileNameStr_result_payment)
                        print(task + '_' + item2 + '=====需要30%预付款')
                        needPay = 1
                    elif payment_infor == '款到发货':
                        ws_sheet0.range('C' + str(29 + offset)).value = '全款，款到发货'
                        ws_sheet0.range('F' + str(29 + offset)).value = ws_sheet0.range('F' + str(18 + offset)).value
                        #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                        my_wbsave(wb, FolderNameStr_result + task + '付款单/' + FileNameStr_result_payment)
                        print(task + '_' + item2 + '=====需要付全款')
                        needPay = 1
                    elif payment_infor == '预付30 %，付清尾款发货':
                        ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                        ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                        #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                        my_wbsave(wb, FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                        print(task + '_' + item2 + '======需要30%预付款，且全款发货')
                        needPay = 1
                    else:
                        print(task + '_' + item2 + '不需要预付款')
                        k = k-1

                    if needPay ==1:
                        wc_sheet0.range('E' + str(6 + k)).value = ws_sheet0.range('F' + str(29 + offset)).value  # 付款金额
                        wc_sheet0.range('B' + str(6 + k)).value = supplier_infor['供应商名称']  # 供应商名称
                        wc_sheet0.range('F' + str(6 + k)).value = ws_sheet0.range('C29').value  # 付款方式
                        if item2_df.shape[0] > 1:
                            #print(str(6 + k))
                            #print(item2_df.shape[0])
                            #print(str(ws_sheet0.range('A8').value))
                            wc_sheet0.range('C' + str(6 + k)).value = str(ws_sheet0.range('A8').value) + '  等' + str(item2_df.shape[0]) + '种产品'
                        else:
                            wc_sheet0.range('C' + str(6 + k)).value = ws_sheet0.range('A8').value
                        wc_sheet0.range('D' + str(6 + k)).value = item2_df.shape[0]
                    j=j+1
                    k=k+1




                wb.close()

        if function ==0:
            #wc.save(FolderNameStr_result + task + '合同/'+ task + '_ContractList.xlsx')
            my_wbsave(wc, FolderNameStr_result + task + '合同/'+ task + '_ContractList.xlsx')
        else:
            #wc.save(FolderNameStr_result + task + '付款单/' + task + '_PaymentList.xlsx')
            my_wbsave(wc, FolderNameStr_result + task + '付款单/' + task + '_PaymentList.xlsx')
        wc.close()
    else:
        print('没有找到{}对应的信息'.format(task))







app.quit()
print('test1')