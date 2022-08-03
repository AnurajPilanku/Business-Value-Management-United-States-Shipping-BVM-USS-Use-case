'''

Author    :  Anuraj Pilanku

Usecase   :  BVM USS Report


'''


import pandas as pd
import sys

path =sys.argv[1]# r"\\acdev01\3M_CAC\anuraj_pilanku"  # sys.argv[1]
constantpath = sys.argv[2]#r"\\acdev01\3M_CAC\BVM_USS\constantData"

ship_date_path = constantpath + "\\" + "ship_date_reasons.xlsx"
Nons_reasons_path = constantpath + "\\" + "Nons_reasons.xlsx"
hold_reasons_path = constantpath + "\\" + "hold_reasons.xlsx"
IOS_errors_path = constantpath + "\\" + "IOS_errors.xlsx"
ship_date_reasons = pd.read_excel(ship_date_path, engine='openpyxl')
Nons_reasons = pd.read_excel(Nons_reasons_path,
                             engine='openpyxl')  # corresponds to Orders_moved_to_NONS_Status_in_USS_due_to_COMS_status in query16.xlsx
hold_reasons = pd.read_excel(hold_reasons_path,
                             engine='openpyxl')  # corresponds to Hold_orders_issue_in_USS in query17.xlsx
IOS_errors = pd.read_excel(IOS_errors_path, engine='openpyxl')  # corresponds to IOS_errored_out_orders in query18.xlsx

path13 = path + "\\" + "query13.xlsx"
data13 = pd.read_excel(path13, engine='openpyxl')
query13_add = pd.merge(data13, ship_date_reasons, on="REV_SHIP_DT_RSN_CD", how="left")

path16 = path + "\\" + "query16.xlsx"
data16 = pd.read_excel(path16, engine='openpyxl')
query16_add = pd.merge(data16, Nons_reasons, on="Orders_moved_to_NONS_Status_in_USS_due_to_COMS_status", how="left")

path17 = path + "\\" + "query17.xlsx"
data17 = pd.read_excel(path17, engine='openpyxl')
query17_add = pd.merge(data17, hold_reasons, on="Hold_orders_issue_in_USS", how="left")

path18 = path + "\\" + "query18.xlsx"
data18 = pd.read_excel(path18, engine='openpyxl')
query18_add = pd.merge(data18, IOS_errors, on="IOS_errored_out_orders", how="left")

# ASN Orders,query1 and query10--Inventory.query1,inventory
path10 = path + "\\" + "query10.xlsx"
path1 = path + "\\" + "query1.xlsx"
inventorypath = constantpath + "\\" + "inventory.xlsx"
data1 = pd.read_excel(path1, engine='openpyxl')
data10 = pd.read_excel(path10, engine='openpyxl')
inventory = pd.read_excel(inventorypath, engine='openpyxl')
ordnum = data10['Order_Number'].tolist()
inven = inventory['LOC_CODE'].tolist()

data1['ASN_Orders'] = ['Y' if m in ordnum else 'N' for m in data1['Order_Number']]

data10['dummy'] = None
data1['inventory'] = ['Y' if n in inven else 'N' for n in data1['LOC_CODE']]
data10[['Order_Number', 'dummy']].to_excel(path + "\\" + "query10_add.xlsx", index=False)
data1[['Order_Number', 'inventory','ASN_Orders']].to_excel(path + "\\" + "inventory_add.xlsx", index=False)

query13_add.to_excel(path + "\\" + "query13_add.xlsx", index=False)
query16_add.to_excel(path + "\\" + "query16_add.xlsx", index=False)
query17_add.to_excel(path + "\\" + "query17_add.xlsx", index=False)
query18_add.to_excel(path + "\\" + "query18_add.xlsx", index=False)
print("success")



