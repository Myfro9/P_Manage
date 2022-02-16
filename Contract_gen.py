import pandas as pd
import os
import datetime
import time
import xlwings as xw
import shutil
import collections
from pandas.api.types import is_datetime64_any_dtype as is_datetime
#import Schedule_gen




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
    #if not os.path.exists(FileNameStr):
    #    os.makedirs(FileNameStr)
    wb_wc.save('./results/temp_result.xlsx')
    shutil.copyfile('./results/temp_result.xlsx',FileNameStr)


def Tech_require(ERP,Techreq_file):  # 读取工艺文件，提取工艺要求
    pd_Techreq = pd.DataFrame(pd.read_excel(Techreq_file))
    sr_Tech_str = pd_Techreq.loc[pd_Techreq['物料编码']==ERP,'工艺要求']
    Tech_str = sr_Tech_str.values
    assert not sr_Tech_str.empty, 'ERP {} 没有工艺要求信息'.format(ERP)
    #print('test')
    return Tech_str

def PurchaseContract_gen(Target_Task,pd_Quote_Infor,pd_Supplier_Infor,
                 FolderNameStr,FileNameStr_ContractList,FileNameStr_PyamentList,
                 FileNameStr_Contract0,FileNameStr_Contract1,FileNameStr_Contract2,
                 FileNameStr_Pyament,FileNameStr_Pyament1,FileNameStr_Pyament2,
                 FolderNameStr_result,
                 TimeStr,VersionCtl,Ver,function,function2):   # funciton: 0: 合同生成； 1：付款单生成 ；function2：0:普通订单； 1：海华订单（新的付款方式，小于2万付全款，否则30%预付，70%尾款）
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False  # 是否实时刷新excel程序的显示内容
    if VersionCtl == 1:
        pd_Quote_Infor = pd_Quote_Infor[pd_Quote_Infor['Ver']==Ver]
    if VersionCtl == 0:
        pre_Ver = '_'
    else:
        pre_Ver = '_' + Ver + '_'
    pd_grouped_Quote = pd_Quote_Infor.groupby('生产计划单号')
###### 按照生产计划单进行分类处理
    for item, item_df in pd_grouped_Quote:
        pd_grouped2_Quote = item_df.groupby('渠道')
        task = item
        if (Target_Task == 'all') or (Target_Task in task):
            if function ==0: #合同生成
                wc = app.books.open(FolderNameStr + FileNameStr_ContractList)   # wc是用于某个生产计划单的合同或付款单信息汇总，wb是用于每一份合同或付款单
            else: # 付款单生成
                wc = app.books.open(FolderNameStr + FileNameStr_PyamentList)
            wc_sheet0 = wc.sheets[0]
            wc_sheet0.range('B43').value = task
            wc_sheet0.range('C3').value = '提交日期： ' + TimeStr


            j =0   # j 是用来计数不同的供应商合同
            k =0   # k 是用来计数不同供应商的付款单
            ###### 在一个生产计划单内，按供应商分类处理
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
                            offset = 50
                        assert item2_df.shape[0]<61 , '采购列表太长，超过60条了！'
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
                            offset = 50
                        assert item2_df.shape[0]<61 , '付款列表太长，超过50条了！'


                    supplier_infor = supplier_infor.iloc[0,:]
                    payment_infor = supplier_infor['账期']

                    sum_total_price =0
                    for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                        ERP = pd_Quote_Infor.loc[item2_df.index[i],'ERP编码']
                        Type = pd_Quote_Infor.loc[item2_df.index[i],'型号']
                        Type2 = pd_Quote_Infor.loc[item2_df.index[i],'规格']
                        Class = pd_Quote_Infor.loc[item2_df.index[i],'供应商类别']
                        num = pd_Quote_Infor.loc[item2_df.index[i],'报价数量']
                        unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                        total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                        sum_total_price = total_price + sum_total_price
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
                                    ws_sheet0.range('C12').value = Type2
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
                                    ws_sheet0.range('C' + str(idx)).value = Type2
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
                                        ws_sheet0.range('H' + str(idx)).value = '预付款后' + str(deliverTime) + '天'
                                    elif payment_infor == '款到发货':
                                        ws_sheet0.range('H' + str(idx)).value = '款到发货'
                                    elif payment_infor == '票到30天':
                                        ws_sheet0.range('H' + str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                    elif payment_infor == '货到付款':
                                        ws_sheet0.range('H' + str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                    elif payment_infor == '预付30 %，付清尾款发货':
                                        ws_sheet0.range('H' + str(idx)).value = '预付款后' + str(deliverTime) + '天'
                                    elif payment_infor == '月结30天':
                                        ws_sheet0.range('H' + str(idx)).value = '合同签订后' + str(deliverTime) + '天'
                                    else:
                                        ws_sheet0.range('H' + str(idx)).value = '合同签订后' + str(deliverTime) + '天'

                                else: # 付款单生成
                                    idx = 8+i
                                    ws_sheet0.range('A'+str(idx)).value = Type
                                    ws_sheet0.range('B'+str(idx)).value = num
                                    ws_sheet0.range('C'+str(idx)).value = unit_price
                                    ws_sheet0.range('E'+str(idx)).value = total_price
                    if function2 == 1:
                        if function ==0 :  # 合同生成
                            for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                                idx = 12 + i
                                Class = pd_Quote_Infor.loc[item2_df.index[i], '供应商类别']
                                # 交期
                                if str(Class) == 'mechanic':
                                    deliverTime = 21
                                else:
                                    deliverTime = 14
                                if sum_total_price < 20000:
                                    ws_sheet0.range('H' + str(idx)).value = '款到发货'
                                elif payment_infor != '款到发货':
                                    ws_sheet0.range('H' + str(idx)).value = '预付款后' + str(deliverTime) + '天'





                    ### 下面对每个合同内的表格做一个统计汇总
                    if function ==0:  # 合同生成
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

                        if function2 == 1:
                            if sum_total_price < 20000:
                                ws_sheet0.range('B' + str(42 + offset)).value = '款到发货'
                                ws_sheet0.range('H12').value = '款到发货'
                                ws_sheet0.range('B' + str(36 + offset)).value = '款到发货'
                            elif payment_infor != '款到发货':
                                ws_sheet0.range('B' + str(42 + offset)).value = '30%预付款，货到票到月结'
                                ws_sheet0.range('H12').value = '预付款后' + str(deliverTime) + '天'
                                ws_sheet0.range('B' + str(36 + offset)).value = '预付款后' + str(deliverTime) + '天发货'


                        if str(Class) == 'nan' :
                            FileNameStr_result =  '_' + task + pre_Ver + item2 + '_合同.xlsx'
                            FileNameStr_result_old =  '_' + task + pre_Ver + item2 + '_合同_old.xlsx'
                        else:
                            FileNameStr_result = Class + '_' + task + pre_Ver + item2 + '_合同.xlsx'
                            FileNameStr_result_old = Class + '_' + task + pre_Ver + item2 + '_合同_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver+ '合同/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver+ '合同/')
                        if os.path.exists(FolderNameStr_result + task + pre_Ver+'合同/' +FileNameStr_result):
                            if os.path.exists(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old):
                                os.remove(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old)
                            os.rename(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result,FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old)
                        #wb.save(FolderNameStr_result + task + '合同/'+ FileNameStr_result)
                        my_wbsave(wb, FolderNameStr_result + task + pre_Ver+ '合同/'+ FileNameStr_result)

                        #### 在合同统计清单中加上一条
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
                            FileNameStr_result_payment =  '_' + task + pre_Ver + item2 + '_付款.xlsx'
                            FileNameStr_result_payment_old =  '_' + task + pre_Ver + item2 + '_付款_old.xlsx'
                        else:
                            FileNameStr_result_payment = Class + '_' + task + pre_Ver + item2 + '_付款.xlsx'
                            FileNameStr_result_payment_old = Class + '_' + task + pre_Ver + item2 + '_付款_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver+'付款单/')
                        if os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/' +FileNameStr_result_payment):
                            if os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old):
                                os.remove(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old)
                            os.rename(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment,FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old)

                        needPay =0

                        if payment_infor == '预付30%+票到30天':
                            ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                            ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver+'付款单/' + FileNameStr_result_payment)
                            print(task + '_' + item2 + '=====需要30%预付款')
                            needPay = 1
                        elif payment_infor == '款到发货':
                            ws_sheet0.range('C' + str(29 + offset)).value = '全款，款到发货'
                            ws_sheet0.range('F' + str(29 + offset)).value = ws_sheet0.range('F' + str(18 + offset)).value
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver +'付款单/' + FileNameStr_result_payment)
                            print(task + '_' + item2 + '=====需要付全款')
                            needPay = 1
                        elif payment_infor == '预付30 %，付清尾款发货':
                            ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                            ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment)
                            print(task + '_' + item2 + '======需要30%预付款，且全款发货')
                            needPay = 1

                        else:
                            print(task + '_' + item2 + '不需要预付款')
                            k = k-1
                        if function2 == 1:
                            if sum_total_price < 20000:
                                ws_sheet0.range('C' + str(29 + offset)).value = '全款，款到发货'
                                ws_sheet0.range('F' + str(29 + offset)).value = ws_sheet0.range(
                                    'F' + str(18 + offset)).value
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要付全款')
                                needPay = 1

                            elif payment_infor != '款到发货':
                                ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要30%预付款')
                                needPay = 1

                        #### 在付款统计清单里加上一条
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

                    # 对于需要尾款支付的项目，再次打开付款模版并生成尾款付款单
                    if function == 1:
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
                            offset = 50
                        assert item2_df.shape[0]<61 , '付款列表太长，超过50条了！'

                        #supplier_infor = supplier_infor.iloc[0, :]
                        #payment_infor = supplier_infor['账期']

                        sum_total_price = 0
                        for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                            ERP = pd_Quote_Infor.loc[item2_df.index[i], 'ERP编码']
                            Type = pd_Quote_Infor.loc[item2_df.index[i], '型号']
                            Type2 = pd_Quote_Infor.loc[item2_df.index[i], '规格']
                            Class = pd_Quote_Infor.loc[item2_df.index[i], '供应商类别']
                            num = pd_Quote_Infor.loc[item2_df.index[i], '报价数量']
                            unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                            total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                            sum_total_price = total_price + sum_total_price
                            if num != None and str(num) != 'nan' and num > 0:
                                if i == 0:
                                    infor0 = supplier_infor['供应商名称']
                                    if str(supplier_infor['开户行']) != 'nan':
                                        infor5 = str(supplier_infor['开户行'])
                                    else:
                                        infor5 = '需补充对方开户行信息'

                                    if str(supplier_infor['账号']) != 'nan':
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

                                    idx = 8 + i
                                    ws_sheet0.range('A' + str(idx)).value = Type
                                    ws_sheet0.range('B' + str(idx)).value = num
                                    ws_sheet0.range('C' + str(idx)).value = unit_price
                                    ws_sheet0.range('E' + str(idx)).value = total_price


                        ### 下面对每个合同内的表格做一个统计汇总

                        if str(Class) == 'nan':
                            FileNameStr_result_payment = '_' + task + pre_Ver + item2 + '_尾款付款.xlsx'
                            FileNameStr_result_payment_old = '_' + task + pre_Ver + item2 + '_尾款付款_old.xlsx'
                        else:
                            FileNameStr_result_payment = Class + '_' + task + pre_Ver + item2 + '_尾款付款.xlsx'
                            FileNameStr_result_payment_old = Class + '_' + task + pre_Ver + item2 + '_尾款付款_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver + '付款单/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver + '付款单/')
                        if os.path.exists(
                                FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment):
                            if os.path.exists(
                                    FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old):
                                os.remove(
                                    FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old)
                            os.rename(FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment,
                                      FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old)

                        needPay = 0
                        if function2 ==0:
                            if payment_infor == '预付30%+票到30天':
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要70%尾款0')
                                needPay = 1

                            elif payment_infor == '预付30 %，付清尾款发货':
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '======需要70%尾款，全款发货')
                                needPay = 1

                        if function2 == 1:
                            if payment_infor != '款到发货' and sum_total_price >= 20000:
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)

                                print(task + '_' + item2 + '=====需要70%尾款1')
                                needPay = 1
                            else:
                                needPay =0

                        #### 在付款统计清单里加上一条
                        '''if needPay == 1:
                            wc_sheet0.range('E' + str(6 + k)).value = ws_sheet0.range(
                                'F' + str(29 + offset)).value  # 付款金额
                            wc_sheet0.range('B' + str(6 + k)).value = supplier_infor['供应商名称']  # 供应商名称
                            wc_sheet0.range('F' + str(6 + k)).value = ws_sheet0.range('C29').value  # 付款方式
                            if item2_df.shape[0] > 1:
                                # print(str(6 + k))
                                # print(item2_df.shape[0])
                                # print(str(ws_sheet0.range('A8').value))
                                wc_sheet0.range('C' + str(6 + k)).value = str(
                                    ws_sheet0.range('A8').value) + '  等' + str(item2_df.shape[0]) + '种产品'
                            else:
                                wc_sheet0.range('C' + str(6 + k)).value = ws_sheet0.range('A8').value
                            wc_sheet0.range('D' + str(6 + k)).value = item2_df.shape[0]
                        j = j + 1
                        k = k + 1'''

                        wb.close()

            ###  存储合同和付款单统计清单
            if function ==0:
                #wc.save(FolderNameStr_result + task + '合同/'+ task + '_ContractList.xlsx')
                if not os.path.exists(FolderNameStr_result + task + pre_Ver + '合同/'):
                    os.makedirs(FolderNameStr_result + task + pre_Ver + '合同/')
                my_wbsave(wc, FolderNameStr_result + task + pre_Ver+'合同/'+ task + pre_Ver + 'ContractList.xlsx')
            else:
                #wc.save(FolderNameStr_result + task + '付款单/' + task + '_PaymentList.xlsx')
                if not os.path.exists(FolderNameStr_result + task + pre_Ver + '付款单/'):
                    os.makedirs(FolderNameStr_result + task + pre_Ver + '付款单/')
                my_wbsave(wc, FolderNameStr_result + task + pre_Ver+'付款单/' + task + pre_Ver+'PaymentList.xlsx')
            wc.close()
        else:
            print('没有找到{}对应的信息'.format(task))

    app.quit()


def OutsourcingContract_gen(Target_Task,pd_Quote_Infor,pd_Supplier_Infor,
                 FolderNameStr,FileNameStr_ContractList,FileNameStr_PyamentList,
                 FileNameStr_OutsourcingContract,
                 FileNameStr_Pyament,
                 FolderNameStr_result,
                 Techreq_file,
                 TimeStr,VersionCtl,Ver,function,function2):   # funciton: 0: 合同生成； 1：付款单生成
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False  # 是否实时刷新excel程序的显示内容
    if VersionCtl == 1:
        pd_Quote_Infor = pd_Quote_Infor[pd_Quote_Infor['Ver']==Ver]
    if VersionCtl == 0:
        pre_Ver = '_'
    else:
        pre_Ver = '_' + Ver + '_'
    pd_grouped_Quote = pd_Quote_Infor.groupby('生产计划单号')
###### 按照生产计划单进行分类处理
    for item, item_df in pd_grouped_Quote:
        pd_grouped2_Quote = item_df.groupby('渠道')
        task = item
        if (Target_Task == 'all') or (Target_Task in task):
            if function ==0: #合同生成
                wc = app.books.open(FolderNameStr + FileNameStr_ContractList)   # wc是用于某个生产计划单的合同或付款单信息汇总，wb是用于每一份合同或付款单
            else: # 付款单生成
                wc = app.books.open(FolderNameStr + FileNameStr_PyamentList)
            wc_sheet0 = wc.sheets[0]
            wc_sheet0.range('B43').value = task
            wc_sheet0.range('C3').value = '提交日期： ' + TimeStr


            j =0   # j 是用来计数不同的供应商合同
            k =0   # k 是用来计数不同供应商的付款单
            ###### 在一个生产计划单内，按供应商分类处理
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
                        assert item2_df.shape[0] < 11+5, '委外加工列表太长，超过15条了！'
                        wb = app.books.open(FolderNameStr + FileNameStr_OutsourcingContract)
                        ws_sheet0 = wb.sheets[0]
                        #ws_sheet1 = wb.sheets[1]
                        offset = 15    # offset 是模版大于10的部分长度
                    else: # 付款单生成
                        wb = app.books.open(FolderNameStr + FileNameStr_Pyament)
                        ws_sheet0 = wb.sheets[0]
                        offset = 15
                    supplier_infor = supplier_infor.iloc[0,:]
                    payment_infor = supplier_infor['账期']
                    sum_total_price = 0

                    for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                        ERP = pd_Quote_Infor.loc[item2_df.index[i],'ERP编码']

                        Type = pd_Quote_Infor.loc[item2_df.index[i],'型号']
                        Class = pd_Quote_Infor.loc[item2_df.index[i],'供应商类别']
                        num = pd_Quote_Infor.loc[item2_df.index[i],'报价数量']
                        unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                        total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                        sum_total_price = total_price + sum_total_price
                        if num != None and str(num) != 'nan' and num > 0:
                            if function == 0:  # 合同生成
                                Tech_str = Tech_require(ERP,Techreq_file)
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
                                    ws_sheet0.range('E'+str(72+5+5)).value = infor1
                                    ws_sheet0.range('B8').value = infor2
                                    ws_sheet0.range('E'+str(74+5+5)).value = infor3
                                    ws_sheet0.range('E'+str(75+5+5)).value = infor4
                                    ws_sheet0.range('E'+str(76+5+5)).value = infor5
                                    ws_sheet0.range('H5').value = task

                                    ws_sheet0.range('B12').value = Type
                                    ws_sheet0.range('B'+str(27+5)).value = Type
                                    ws_sheet0.range('D12').value = ERP
                                    ws_sheet0.range('C'+str(27+5)).value = ERP
                                    ws_sheet0.range('D'+str(27+5)).value = Tech_str
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
                                    ws_sheet0.range('B' + str(idx+15+5)).value = Type

                                    ws_sheet0.range('D'+str(idx)).value = ERP
                                    ws_sheet0.range('C' + str(idx+15+5)).value = ERP
                                    ws_sheet0.range('D' + str(idx+15+5)).value = Tech_str

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
                    if function2 == 1:
                        if function == 0:  # 合同生成
                            for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                                idx = 12 + i
                                Class = pd_Quote_Infor.loc[item2_df.index[i], '供应商类别']
                                # 交期
                                if str(Class) == 'mechanic':
                                    deliverTime = 21
                                else:
                                    deliverTime = 14
                                if sum_total_price < 20000:
                                    ws_sheet0.range('H' + str(idx)).value = '款到发货'
                                elif payment_infor != '款到发货':
                                    ws_sheet0.range('H' + str(idx)).value = '预付款后' + str(deliverTime) + '天'

                    ### 下面对每个合同内的表格做一个统计汇总
                    if function ==0:  # 合同生成
                        sumed_price =ws_sheet0.range('G'+str(22+5)).value
                        sumed_price_Chinese = int2Chnese(int(sumed_price))
                        ws_sheet0.range('A'+str(22+5+5)).value = '以上单价含13%增值税，大写金额：'+ sumed_price_Chinese + '元整'
                        # 交期
                        if str(Class) == 'mechanic':
                            deliverTime = 21
                        else:
                            deliverTime =14
                        # 帐期
                        if payment_infor == '预付30%+票到30天':
                            ws_sheet0.range('B'+str(54+5)).value = '30%预付款，货到票到30天月结'
                            ws_sheet0.range('H12').value = '预付款后'+ str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '预付款后'+ str(deliverTime) + '天发货'
                        elif payment_infor == '款到发货':
                            ws_sheet0.range('B'+str(54+5)).value = '款到发货'
                            ws_sheet0.range('H12').value = '款到发货'
                            ws_sheet0.range('B' + str(48 + 5)).value = '款到发货'
                        elif payment_infor == '票到30天':
                            ws_sheet0.range('B'+str(54+5)).value = '货到票到30天月结'
                            ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '合同签订后' + str(deliverTime) + '天发货'
                        elif payment_infor == '货到付款':
                            ws_sheet0.range('B'+str(54+5)).value = '货到付款'
                            ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '合同签订后' + str(deliverTime) + '天发货'
                        elif payment_infor == '预付30 %，付清尾款发货':
                            ws_sheet0.range('B' + str(54 + 5)).value = '预付30 %，付清尾款发货'
                            ws_sheet0.range('H12').value = '预付款后' + str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '预付款后' + str(deliverTime) + '天发货'
                        elif payment_infor == '月结30天':
                            ws_sheet0.range('B' + str(54 + 5)).value = '货到月结30天'
                            ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '合同签订后' + str(deliverTime) + '天发货'
                        else:
                            ws_sheet0.range('B'+str(54+5)).value = '货到月结'
                            ws_sheet0.range('H12').value = '合同签订后' + str(deliverTime) + '天'
                            ws_sheet0.range('B' + str(48 + 5)).value = '合同签订后' + str(deliverTime) + '天发货'

                        if function2 == 1:
                            if sum_total_price < 20000:
                                ws_sheet0.range('B' + str(42 + offset)).value = '款到发货'
                                ws_sheet0.range('H12').value = '款到发货'
                                ws_sheet0.range('B' + str(36 + offset)).value = '款到发货'
                            elif payment_infor != '款到发货':
                                ws_sheet0.range('B' + str(42 + offset)).value = '30%预付款，货到票到月结'
                                ws_sheet0.range('H12').value = '预付款后' + str(deliverTime) + '天'
                                ws_sheet0.range('B' + str(36 + offset)).value = '预付款后' + str(deliverTime) + '天发货'


                        if str(Class) == 'nan' :
                            FileNameStr_result =  '_' + task + pre_Ver + item2 + '_委外合同.xlsx'
                            FileNameStr_result_old =  '_' + task + pre_Ver + item2 + '_委外合同_old.xlsx'
                        else:
                            FileNameStr_result = Class + '_' + task + pre_Ver + item2 + '_委外合同.xlsx'
                            FileNameStr_result_old = Class + '_' + task + pre_Ver + item2 + '_委外合同_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver+ '合同/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver+ '合同/')
                        if os.path.exists(FolderNameStr_result + task + pre_Ver+'合同/' +FileNameStr_result):
                            if os.path.exists(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old):
                                os.remove(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old)
                            os.rename(FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result,FolderNameStr_result + task + pre_Ver+'合同/'+ FileNameStr_result_old)
                        #wb.save(FolderNameStr_result + task + '合同/'+ FileNameStr_result)
                        my_wbsave(wb, FolderNameStr_result + task + pre_Ver+ '合同/'+ FileNameStr_result)

                        #### 在合同统计清单中加上一条
                        wc_sheet0.range('E'+str(6+j)).value = ws_sheet0.range('G'+str(22+5)).value  #合同金额
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
                            FileNameStr_result_payment =  '_' + task + pre_Ver + item2 + '_委外付款.xlsx'
                            FileNameStr_result_payment_old =  '_' + task + pre_Ver + item2 + '_委外付款_old.xlsx'
                        else:
                            FileNameStr_result_payment = Class + '_' + task + pre_Ver + item2 + '_委外付款.xlsx'
                            FileNameStr_result_payment_old = Class + '_' + task + pre_Ver + item2 + '_委外付款_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver+'付款单/')
                        if os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/' +FileNameStr_result_payment):
                            if os.path.exists(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old):
                                os.remove(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old)
                            os.rename(FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment,FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment_old)

                        needPay =0

                        if payment_infor == '预付30%+票到30天':
                            ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                            ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver+'付款单/' + FileNameStr_result_payment)
                            print(task + '_' + item2 + '=====需要30%预付款')
                            needPay = 1
                        elif payment_infor == '款到发货':
                            ws_sheet0.range('C' + str(29 + offset)).value = '全款，款到发货'
                            ws_sheet0.range('F' + str(29 + offset)).value = ws_sheet0.range('F' + str(18 + offset)).value
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver +'付款单/' + FileNameStr_result_payment)
                            print(task + '_' + item2 + '=====需要付全款')
                            needPay = 1
                        elif payment_infor == '预付30 %，付清尾款发货':
                            ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                            ws_sheet0.range('F' + str(29 + offset)).value = int(ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                            #wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                            my_wbsave(wb, FolderNameStr_result + task + pre_Ver+'付款单/'+ FileNameStr_result_payment)
                            print(task + '_' + item2 + '======需要30%预付款，且全款发货')
                            needPay = 1
                        else:
                            print(task + '_' + item2 + '不需要预付款')
                            k = k-1

                        if function2 == 1:
                            if sum_total_price < 20000:
                                ws_sheet0.range('C' + str(29 + offset)).value = '全款，款到发货'
                                ws_sheet0.range('F' + str(29 + offset)).value = ws_sheet0.range(
                                    'F' + str(18 + offset)).value
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要付全款')
                                needPay = 1

                            elif payment_infor != '款到发货':
                                ws_sheet0.range('C' + str(29 + offset)).value = '30%预付款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.3
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要30%预付款')
                                needPay = 1
                        #### 在付款统计清单里加上一条
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

                    # 对于需要尾款支付的项目，再次打开付款模版并生成尾款付款单
                    if function == 1:
                        wb = app.books.open(FolderNameStr + FileNameStr_Pyament)
                        ws_sheet0 = wb.sheets[0]
                        offset = 15

                        # supplier_infor = supplier_infor.iloc[0, :]
                        # payment_infor = supplier_infor['账期']

                        sum_total_price = 0
                        for i in range(item2_df.shape[0]):  ## for 每个生产计划单的每一个供应商对应的多个物料：
                            ERP = pd_Quote_Infor.loc[item2_df.index[i], 'ERP编码']
                            Type = pd_Quote_Infor.loc[item2_df.index[i], '型号']
                            Type2 = pd_Quote_Infor.loc[item2_df.index[i], '规格']
                            Class = pd_Quote_Infor.loc[item2_df.index[i], '供应商类别']
                            num = pd_Quote_Infor.loc[item2_df.index[i], '报价数量']
                            unit_price = pd_Quote_Infor.loc[item2_df.index[i], '单价（元）']
                            total_price = pd_Quote_Infor.loc[item2_df.index[i], '小计']
                            sum_total_price = total_price + sum_total_price
                            if num != None and str(num) != 'nan' and num > 0:
                                if i == 0:
                                    infor0 = supplier_infor['供应商名称']
                                    if str(supplier_infor['开户行']) != 'nan':
                                        infor5 = str(supplier_infor['开户行'])
                                    else:
                                        infor5 = '需补充对方开户行信息'

                                    if str(supplier_infor['账号']) != 'nan':
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

                                    idx = 8 + i
                                    ws_sheet0.range('A' + str(idx)).value = Type
                                    ws_sheet0.range('B' + str(idx)).value = num
                                    ws_sheet0.range('C' + str(idx)).value = unit_price
                                    ws_sheet0.range('E' + str(idx)).value = total_price

                        ### 下面对每个合同内的表格做一个统计汇总

                        if str(Class) == 'nan':
                            FileNameStr_result_payment = '_' + task + pre_Ver + item2 + '_尾款付款.xlsx'
                            FileNameStr_result_payment_old = '_' + task + pre_Ver + item2 + '_尾款付款_old.xlsx'
                        else:
                            FileNameStr_result_payment = Class + '_' + task + pre_Ver + item2 + '_尾款付款.xlsx'
                            FileNameStr_result_payment_old = Class + '_' + task + pre_Ver + item2 + '_尾款付款_old.xlsx'

                        if not os.path.exists(FolderNameStr_result + task + pre_Ver + '付款单/'):
                            os.makedirs(FolderNameStr_result + task + pre_Ver + '付款单/')
                        if os.path.exists(
                                FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment):
                            if os.path.exists(
                                    FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old):
                                os.remove(
                                    FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old)
                            os.rename(FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment,
                                      FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment_old)

                        needPay = 0
                        if function2 == 0:
                            if payment_infor == '预付30%+票到30天':
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '=====需要70%尾款0')
                                needPay = 1

                            elif payment_infor == '预付30 %，付清尾款发货':
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)
                                print(task + '_' + item2 + '======需要70%尾款，全款发货')
                                needPay = 1

                        if function2 == 1:
                            if payment_infor != '款到发货' and sum_total_price >= 20000:
                                ws_sheet0.range('C' + str(29 + offset)).value = '70%尾款'
                                ws_sheet0.range('F' + str(29 + offset)).value = int(
                                    ws_sheet0.range('F' + str(18 + offset)).value) * 0.7
                                # wb.save(FolderNameStr_result + task + '付款单/'+ FileNameStr_result_payment)
                                my_wbsave(wb,
                                          FolderNameStr_result + task + pre_Ver + '付款单/' + FileNameStr_result_payment)

                                print(task + '_' + item2 + '=====需要70%尾款1')
                                needPay = 1
                            else:
                                needPay = 0

                        #### 在付款统计清单里加上一条
                        '''if needPay == 1:
                            wc_sheet0.range('E' + str(6 + k)).value = ws_sheet0.range(
                                'F' + str(29 + offset)).value  # 付款金额
                            wc_sheet0.range('B' + str(6 + k)).value = supplier_infor['供应商名称']  # 供应商名称
                            wc_sheet0.range('F' + str(6 + k)).value = ws_sheet0.range('C29').value  # 付款方式
                            if item2_df.shape[0] > 1:
                                # print(str(6 + k))
                                # print(item2_df.shape[0])
                                # print(str(ws_sheet0.range('A8').value))
                                wc_sheet0.range('C' + str(6 + k)).value = str(
                                    ws_sheet0.range('A8').value) + '  等' + str(item2_df.shape[0]) + '种产品'
                            else:
                                wc_sheet0.range('C' + str(6 + k)).value = ws_sheet0.range('A8').value
                            wc_sheet0.range('D' + str(6 + k)).value = item2_df.shape[0]
                        j = j + 1
                        k = k + 1'''

                        wb.close()


            ###  存储合同和付款单统计清单
            if function ==0:
                #wc.save(FolderNameStr_result + task + '合同/'+ task + '_ContractList.xlsx')
                my_wbsave(wc, FolderNameStr_result + task + pre_Ver+'合同/'+ task + pre_Ver + '委外_ContractList.xlsx')
            else:
                #wc.save(FolderNameStr_result + task + '付款单/' + task + '_PaymentList.xlsx')
                my_wbsave(wc, FolderNameStr_result + task + pre_Ver+'付款单/' + task + pre_Ver+'委外_PaymentList.xlsx')
            wc.close()
        else:
            print('没有找到{}对应的信息'.format(task))

    app.quit()

def Quote_infor_Clasify(pd_Quote_Infor):  # 区分是采购询价信息还是加工询价信息
    pd_Quote_Infor_purchase = pd_Quote_Infor[~pd_Quote_Infor['ERP编码'].str.contains('H').fillna(False)]
    pd_Quote_Infor_purchase.index = range(pd_Quote_Infor_purchase.shape[0])
    pd_Quote_Outsourcing = pd_Quote_Infor[pd_Quote_Infor['ERP编码'].str.contains('H').fillna(False)]
    pd_Quote_Outsourcing.index = range(pd_Quote_Outsourcing.shape[0])
    #print('test')
    return pd_Quote_Infor_purchase,pd_Quote_Outsourcing

def Clean_Quote_info(pd_QuoteInfo,pd_Supplier_Infor):
    pd_QuoteInfo.index = range(pd_QuoteInfo.shape[0])
    idx = pd_QuoteInfo.index
    for i in range(pd_QuoteInfo.shape[0]):
        ERP = pd_QuoteInfo.loc[idx[i], 'ERP编码']
        Task = pd_QuoteInfo.loc[idx[i], '生产计划单号']
        Supplier = pd_QuoteInfo.loc[idx[i], '渠道']
        if str(Task) in str(Supplier):
            #print('Warning # 0: 生产任务单 ', Task,' 中 ',ERP, ' 的供应商是自己，死循环！！')
            pd_QuoteInfo.drop(index=idx[i],axis=0,inplace=True)
    pd_QuoteInfo.index = range(pd_QuoteInfo.shape[0])
    key_list = ['RW', 'YF', 'GC']
    for i in range(pd_QuoteInfo.shape[0]):
        ERP = pd_QuoteInfo.loc[idx[i], 'ERP编码']
        Task = pd_QuoteInfo.loc[idx[i], '生产计划单号']
        Supplier = str(pd_QuoteInfo.loc[idx[i], '渠道'])
        if any(key in str(Supplier) for key in key_list):   ## 当合并采购任务到其它生产计划单的情况，确保其它生产计划中有该ERP编码的采购信息
            Task_others = pd_QuoteInfo.loc[pd_QuoteInfo.index[i],'渠道']
            Target =  pd_QuoteInfo[(pd_QuoteInfo['生产计划单号'] == Task_others) & (pd_QuoteInfo['ERP编码'] == ERP)]
            if Target.empty:
                print('Error #3-1 : 在生产计划单',Task,' 和',Task_others,' 中，都没有ERP编码：',ERP,' 的采购信息！\n' )
                pd_QuoteInfo.loc[idx[i], '渠道']='nan'
        else:
            if not (('苏州全波' in Supplier) or (Supplier == 'nan') or ('特殊采购' in Supplier)):
                pd_Supplier_Target = pd_Supplier_Infor[pd_Supplier_Infor['供应商名称']==Supplier]
                if pd_Supplier_Target.empty:
                    #print('Error #4 : 供应商档案中没有 ',Supplier,' 的信息！\n')
                    pd_QuoteInfo.loc[idx[i], '渠道']='nan'

    pd_QuoteInfo.index = range(pd_QuoteInfo.shape[0])
    return pd_QuoteInfo

def Check_quote_info(pd_Quote_Infor,pd_Supplier_Infor):
    #pd_Quote_Infor = Schedule_gen.Clean_Quote_info(pd_Quote_Infor)
    f = open('Quote_errlog.txt', 'w')
    Err =0

    ## 检查询价信息表
    key_list = ['RW', 'YF', 'GC']
    clm = pd_Quote_Infor.columns
    #print(clm[1:11].values)
    if pd_Quote_Infor.shape[1] < 11 or collections.Counter(clm[1:11].values) != collections.Counter(['生产计划单号','ERP编码','型号','需求数量','供应商类别','报价数量','单价（元）','小计','渠道','Ver']):
        Err =1
        print('Error #1 : 询价结果的列格式不对\n',file=f)
    for i in range(pd_Quote_Infor.shape[0]):
        Task = pd_Quote_Infor.loc[pd_Quote_Infor.index[i],'生产计划单号']
        ERP = pd_Quote_Infor.loc[pd_Quote_Infor.index[i], 'ERP编码']
        Supplier = str(pd_Quote_Infor.loc[pd_Quote_Infor.index[i], '渠道'])
        if Task == pd_Quote_Infor.loc[pd_Quote_Infor.index[i],'渠道']:   ## 避免合并采购的生产任务单号指向自己
            print('Error #2 : 对于',ERP,'生产计划单',Task,'的供应商指向自己，死循环\n',file=f)
            Err =1

        elif any(key in str(Supplier) for key in key_list):   ## 当合并采购任务到其它生产计划单的情况，确保其它生产计划中有该ERP编码的采购信息
            Task_others = pd_Quote_Infor.loc[pd_Quote_Infor.index[i],'渠道']

            Target =  pd_Quote_Infor[(pd_Quote_Infor['生产计划单号'] == Task_others) & (pd_Quote_Infor['ERP编码'] == ERP)]
            if Target.empty:
                print('Error #3-1 : 在生产计划单',Task,' 和',Task_others,' 中，都没有ERP编码：',ERP,' 的采购信息！\n',file=f )
                Err =1
            else:
                if Target.shape[0]>1:
                    Target_sr = Target['Ver'].value_counts()
                    #print(len(Target_sr))
                    if len(Target_sr) != Target.shape[0]:
                        print('Error #3-2 : 在生产计划单',Task_others,' 中，有多条ERP编码：',ERP,' 的采购信息！\n',file=f)
                        Err =1
                else:
                    if any(key in Target['渠道'] for key in key_list):
                        print('Error #3-3 : 在生产计划单', Task, ' 的合并采购生产计划单', Task_others, ' 中，', ERP, '仍然指向其它的生产任务单！\n',file=f)
                        Err = 1
        else:
            if not (('苏州全波' in Supplier) or (Supplier == 'nan') or ('特殊采购' in Supplier)):
                pd_Supplier_Target = pd_Supplier_Infor[pd_Supplier_Infor['供应商名称']==Supplier]
                if pd_Supplier_Target.empty:
                    print('Error #4 : 供应商档案中没有 ',Supplier,' 的信息！\n',file=f)
                    Err =1

    ## 下面检查供应商信息表
    Item_Counts = pd_Supplier_Infor['供应商名称'].value_counts()
    pd_Supplier_Counts = pd.DataFrame({'col1': Item_Counts.index, 'col2':Item_Counts.values})
    pd_Supplier_Counts = pd_Supplier_Counts[pd_Supplier_Counts['col2']>1]
    payment_key_list = ['票到30天', '月结30天', '款到发货', '预付30%+票到30天', '货到付款', '预付50%货款，付清尾款发货。', '预付30%，付清尾款发货',
                        '预付50%货款，尾款根据项目背靠背支付。', '预付50%，尾款月结30天','其他']
    if not pd_Supplier_Counts.empty:
        for i in range(pd_Supplier_Counts.shape[0]):
            print('Error #5 : 供应商档案中，', pd_Supplier_Counts.iloc[i,0],'有',pd_Supplier_Counts.iloc[i,1],'条信息 \n',file=f)
            Err =1

    for i in range(pd_Supplier_Infor.shape[0]):
        Supplier = str(pd_Supplier_Infor.loc[pd_Supplier_Infor.index[i],'供应商名称']).strip()
        Payment_Infor = str(pd_Supplier_Infor.loc[pd_Supplier_Infor.index[i],'账期']).strip()
        if not any(key == Payment_Infor for key in payment_key_list):
            if Payment_Infor != 'nan':
                print('Error #6 : 供应商档案中，',Supplier, ' 的付款方式为：' , Payment_Infor, ' ，不符合规定！\n',file=f)

    sr_Contract_time = pd.to_datetime(pd_Quote_Infor['盖章日期'])
    if not is_datetime(sr_Contract_time):
        Err = 1
        print('Error #7 : 合同盖章日期应该为日期格式 \n',file=f)
    return Err


def main():
    function2 = 1  # function2 : 0, 普通订单； 1：海华订单
    ####### 导入询价结果
    FolderNameStr = './Purchase_Rawdata/23询价结果/'
    Quote_result = '合并汇总表-版本45-20220216A'

    FileNameStr = Quote_result + '.xlsx'
    pd_Quote_Infor = pd.DataFrame(pd.read_excel(FolderNameStr + FileNameStr, sheet_name='操作'))
    FolderNameStr_result = './Results/' + Quote_result + '/'
    if not os.path.exists(FolderNameStr_result):
        os.makedirs(FolderNameStr_result)

    ####### 导入供应商信息
    FolderNameStr = './Purchase_Rawdata/22供应商档案/'
    FileNameStr = '供应商档案 38-20220216.xlsx'
    pd_Supplier_Infor = pd.DataFrame(pd.read_excel(FolderNameStr + FileNameStr))

    ####### 导入工艺文件
    FolderNameStr0 = './Purchase_Rawdata/30工艺要求/'
    FileNameStr0 = '焊接工艺要求20220216.xls'

    ####### 打开合同模版
    FolderNameStr = './Purchase_Rawdata/'
    FileNameStr_Contract0 = 'Excel_templates/Contract Template.xlsx'
    FileNameStr_Contract1 = 'Excel_templates/Contract Template_long.xlsx'
    FileNameStr_Contract2 = 'Excel_templates/Contract Template_longest.xlsx'
    FileNameStr_OutsourcingContract = 'Excel_templates/委托加工合同模版.xlsx'
    FileNameStr1 = 'Excel_templates/Contract target.xlsx'

    ####### 打开付款模版
    FileNameStr_Pyament = 'Excel_templates/付款单模板.xlsx'
    FileNameStr_Pyament1 = 'Excel_templates/付款单模板_long.xlsx'
    FileNameStr_Pyament2 = 'Excel_templates/付款单模板_longest.xlsx'

    ####### 合同汇总列表
    FileNameStr_ContractList = 'Excel_templates/合同申请单模版.xlsx'
    FileNameStr_PyamentList = 'Excel_templates/预付款申请单模版.xlsx'

    ####### 合同存储

    TimeStr = time.strftime("%Y-%m-%d", time.localtime(time.time()))

    pd_Quote_Infor = Clean_Quote_info(pd_Quote_Infor,pd_Supplier_Infor)
    Err = Check_quote_info(pd_Quote_Infor,pd_Supplier_Infor)

    #Tasklist = ['RW202107-C','RW202107-D','RW202108-B-4','RW202111-A','RW202111-B','RW202111-C','RW202111-E','RW202111-F','RW202111-G','RW202111-H']
    #Tasklist = ['RW202111-I-1']
    #Tasklist = ['YF202110-A','RW202107-B-1','RW202108-A-1','RW202108-B-4','RW202108-D-1','RW202111-I','RW202111-A']
    Tasklist = [
           'RW202111-I',
           'RW202111-L',
           'RW202201-A-1',
           'RW202201-A-4',
           'RW202201-A-5',
           'RW202201-A-6',
           'RW202201-A-7',
           'RW202201-A-8',
           'RW202201-B-1',
           'RW202201-B-2',
           'RW202201-B-3',
           'RW202201-B-4',
           'RW202201-C-6',
           'RW202201-C-7',
           'RW202201-D'
        ]

    Verlist = [
        10,
         6,
        11,
         6,
         7,
         1,
         2,
         6,
         4,
         3,
         4,
         1,
         3,
         3,
         4

   ]    # 0: 没有Version
    '''Tasklist = [
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',
                #'RW202201-A-1',

                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',
                #'RW202201-A-2',

                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',
                #'RW202201-A-3',

                #'RW202201-A-4',
                #'RW202201-A-4',
                #'RW202201-A-4',
                #'RW202201-A-4',
                #'RW202201-A-4',
                #'RW202201-A-4',

                #'RW202201-A-5',
                #'RW202201-A-5',
                #'RW202201-A-5',
                #'RW202201-A-5',
                #'RW202201-A-5',
                #'RW202201-A-5',

                #'RW202201-A-6',
                #
                #'RW202201-A-7',
                #'RW202201-A-7',
                #
                #'RW202201-A-8',
                #'RW202201-A-8',
                #'RW202201-A-8',
                #'RW202201-A-8',
                #'RW202201-A-8',
                #
                #'RW202201-B-1',
                #'RW202201-B-1',
                #'RW202201-B-1',
                #
                #'RW202201-B-2',
                #'RW202201-B-2',
                #
                #'RW202201-B-3',
                #'RW202201-B-3',
                #'RW202201-B-3',
                #
                #'RW202201-C-1',
                #'RW202201-C-1',
                #'RW202201-C-1',
                #
                #'RW202201-C-2',
                #'RW202201-C-2',
                #'RW202201-C-2',
                #
                #'RW202201-C-3',
                #'RW202201-C-3',
                #
                #'RW202201-C-4',
                #
                #'RW202201-C-5',
                #'RW202201-C-5',
                #
                #'RW202201-C-6',
                #'RW202201-C-6',
                #'RW202201-C-6',
                #
                #'RW202201-C-7',
                #'RW202201-C-7',
                #'RW202201-C-7',
                
                'RW202201-D',
                'RW202201-D',
                'RW202201-D' 
                ]
    Verlist = [#0,1,2,3,4,5,6,7,8,9,
               #0,1,2,3,4,5,6,7,8,
               #0,1,2,3,4,5,6,7,
               #0,1,2,3,4,5,
               #0,1,2,3,4,5,6,
               #0,                 #A6
               #0,1,               #A7
               #0,1,2,3,4,         #A8
               #
               #0,1,2, #B1
               #0,1,   #B2
               #0,1,2, #B3
               #0,1,2, #C1
               #0,1,2, #C2
               #0,1,       #C3
               #0,      #C4
               #0,1,   #C5
               #0,1,2, #C6
               #0,1,2, #C7
               0,1,2
              ]'''
   

    for i in range(len(Tasklist)):

        if Verlist[i] == -1:
            VersionCtl = 0  # 0: 生成整个项目的合同； 1：之生成对应的Ver
        else:
            VersionCtl = 1
        Ver = 'Ver'+str(Verlist[i])

        Target_Task = Tasklist[i]  # 'all' 或者特定生产计划单号,可以包含关系


        if Err ==1:
            print('询价信息表 或 供应商档案 有错误，具体查阅Quote_errlog.txt !')
        else:
            pd_Quote_Infor_purchase,pd_Quote_Infor_Outsourcing = Quote_infor_Clasify(pd_Quote_Infor)  # 将询价结果拆分成采购和委外加工

            if not pd_Quote_Infor_purchase.empty:
                function = 0  # 采购合同生成
                PurchaseContract_gen(Target_Task,pd_Quote_Infor_purchase,pd_Supplier_Infor,
                             FolderNameStr, FileNameStr_ContractList, FileNameStr_PyamentList,
                             FileNameStr_Contract0, FileNameStr_Contract1, FileNameStr_Contract2,
                             FileNameStr_Pyament, FileNameStr_Pyament1, FileNameStr_Pyament2,
                             FolderNameStr_result,
                             TimeStr, VersionCtl, Ver, function,function2)
                function = 1  # 采购付款单生成
                PurchaseContract_gen(Target_Task, pd_Quote_Infor_purchase, pd_Supplier_Infor,
                             FolderNameStr, FileNameStr_ContractList, FileNameStr_PyamentList,
                             FileNameStr_Contract0, FileNameStr_Contract1, FileNameStr_Contract2,
                             FileNameStr_Pyament, FileNameStr_Pyament1, FileNameStr_Pyament2,
                             FolderNameStr_result,
                             TimeStr, VersionCtl, Ver, function,function2)

            if not pd_Quote_Infor_Outsourcing.empty:
                function = 0  # 委外合同生成
                OutsourcingContract_gen(Target_Task, pd_Quote_Infor_Outsourcing, pd_Supplier_Infor,
                                        FolderNameStr, FileNameStr_ContractList, FileNameStr_PyamentList,
                                        FileNameStr_OutsourcingContract,
                                        FileNameStr_Pyament1,
                                        FolderNameStr_result,
                                        FolderNameStr0+FileNameStr0,
                                        TimeStr, VersionCtl, Ver, function,function2)

                function = 1  # 委外付款单生成
                OutsourcingContract_gen(Target_Task, pd_Quote_Infor_Outsourcing, pd_Supplier_Infor,
                                        FolderNameStr, FileNameStr_ContractList, FileNameStr_PyamentList,
                                        FileNameStr_OutsourcingContract,
                                        FileNameStr_Pyament1,
                                        FolderNameStr_result,
                                        FolderNameStr0 + FileNameStr0,
                                        TimeStr, VersionCtl, Ver, function,function2)

if __name__ == "__main__":
        main()