import pandas as pd

#######
# FileNameStr_PrjList: The file name of the on-going project list
# FolderNameStr_BOM: The folder includes all BOM information
#  The result pd_BOMofALLPrj contains BOM list for all project together

# pd_BOMofALLPrj = BOM_of_ALLPrjGen(FileNameStr_PrjList,FolderNameStr_BOM)


#######
# The file name of stock information

#######
# FileNameStr_AllContract: The file name of all history purchase contract record
# pd_OnGoingContract = OnGoing_Prj_ContractGen(FileNameStr_PrjList, FileNameStr_AllContract)


#######
# FileNameStr_Outstock: The file name of all history Out stock record
# pd_OnGoingOutstock = OnGoing_OutStockGen(FileNameStr_PrjList, FileNameStr_Outstock)

#######
# FileNameStr_Instock: The file name of all history In stock record
# pd_OnGoingInstock = OnGoing_OutStockGen(pd_OnGoingContract, FileNameStr_Instock)