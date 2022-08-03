'''

Author     :  AnurajPilanku

usecase    :  BVM USS Report

Code Utility : Extracting data into single sheet

'''


import pandas as pd
import sys

path=sys.argv[1]#r"\\acdev01\3M_CAC\anuraj_pilanku"#sys.argv[1]
constantpath=sys.argv[2]#r"\\acdev01\3M_CAC\BVM_USS\constantData"#sys.argv[2]

path1=path+"\\"+"query1.xlsx"
path2=path+"\\"+"query2.xlsx"
path3=path+"\\"+"query3.xlsx"
path4=path+"\\"+"query4.xlsx"
path5=path+"\\"+"query5.xlsx"
path6=path+"\\"+"query6.xlsx"
path7=path+"\\"+"query7.xlsx"
path8=path+"\\"+"query8.xlsx"
path9=path+"\\"+"query9.xlsx"
#path10=path+"\\"+"query10.xlsx"
path10=path+"\\"+"query10_add.xlsx"
path11=path+"\\"+"query11.xlsx"
path12=path+"\\"+"query12.xlsx"
path13=path+"\\"+"query13_add.xlsx"
path14=path+"\\"+"query14.xlsx"
path15=path+"\\"+"query15.xlsx"
path16=path+"\\"+"query16_add.xlsx"
path17=path+"\\"+"query17_add.xlsx"
path18=path+"\\"+"query18_add.xlsx"
path19=path+"\\"+"query19.xlsx"
inventorypath=path+"\\"+"inventory_add.xlsx"


#ship_date_path=constantpath+"\\"+"ship_date_reasons.xlsx"
#Nons_reasons_path=constantpath+"\\"+"Nons_reasons.xlsx"
#hold_reasons_path=constantpath+"\\"+"hold_reasons.xlsx"
#IOS_errors_path=constantpath+"\\"+"IOS_errors.xlsx"


outputpath=r"\\acdev01\3M_CAC\anuraj_pilanku\jin.xlsx"#sys.argv[3]
commoncolumnname="Order_Number"
data1=pd.read_excel(path1,engine='openpyxl')
#data2=data2.drop(['LOC_CODE'],axis=1)
#data1['Order_Number'].str.strip()#removing extra white spaces from values in a column
data2=pd.read_excel(path2,engine='openpyxl')
data3=pd.read_excel(path3,engine='openpyxl')
data4=pd.read_excel(path4,engine='openpyxl')
data5=pd.read_excel(path5,engine='openpyxl')
data6=pd.read_excel(path6,engine='openpyxl')
data7=pd.read_excel(path7,engine='openpyxl')
data8=pd.read_excel(path8,engine='openpyxl')
data9=pd.read_excel(path9,engine='openpyxl')
data10=pd.read_excel(path10,engine='openpyxl')
data11=pd.read_excel(path11,engine='openpyxl')
data12=pd.read_excel(path12,engine='openpyxl')
data13=pd.read_excel(path13,engine='openpyxl')
data14=pd.read_excel(path14,engine='openpyxl')
data15=pd.read_excel(path15,engine='openpyxl')
data16=pd.read_excel(path16,engine='openpyxl')
data17=pd.read_excel(path17,engine='openpyxl')
data18=pd.read_excel(path18,engine='openpyxl')
data19=pd.read_excel(path19,engine='openpyxl')
inventory=pd.read_excel(inventorypath,engine='openpyxl')

#STATIC DATA
#ship_date_reasons=pd.read_excel(ship_date_path,engine='openpyxl')#corresponds to REV_SHIP_DT_RSN_CD in query13.xlsx
#Nons_reasons=pd.read_excel(Nons_reasons_path,engine='openpyxl')#corresponds to Orders_moved_to_NONS_Status_in_USS_due_to_COMS_status in query16.xlsx
#hold_reasons=pd.read_excel(hold_reasons_path,engine='openpyxl')#corresponds to Hold_orders_issue_in_USS in query17.xlsx
#IOS_errors=pd.read_excel(IOS_errors_path,engine='openpyxl')#corresponds to IOS_errored_out_orders in query18.xlsx

join1=pd.merge(data1,data2,on=commoncolumnname,how="left")
join2=pd.merge(join1,data3,on=commoncolumnname,how="left")
join3=pd.merge(join2,data4,on=commoncolumnname,how="left")
join4=pd.merge(join3,data5,on=commoncolumnname,how="left")
join5=pd.merge(join4,data6,on=commoncolumnname,how="left")
join6=pd.merge(join5,data7,on=commoncolumnname,how="left")
join7=pd.merge(join6,data8,on=commoncolumnname,how="left")
join8=pd.merge(join7,data9,on=commoncolumnname,how="left")
join9=pd.merge(join8,data10,on=commoncolumnname,how="left")
join10=pd.merge(join9,data11,on=commoncolumnname,how="left")
join11=pd.merge(join10,data12,on=commoncolumnname,how="left")
join12=pd.merge(join11,data13,on=commoncolumnname,how="left")
join13=pd.merge(join12,data14,on=commoncolumnname,how="left")
join14=pd.merge(join13,data15,on=commoncolumnname,how="left")
join15=pd.merge(join14,data16,on=commoncolumnname,how="left")
join16=pd.merge(join15,data17,on=commoncolumnname,how="left")
join17=pd.merge(join16,data18,on=commoncolumnname,how="left")
join18=pd.merge(join17,data19,on=commoncolumnname,how="left")
join19=pd.merge(join18,inventory,on=commoncolumnname,how="left")

#STATIC DATA
#join19=pd.merge(join18,ship_date_reasons,on="REV_SHIP_DT_RSN_CD",how="left")
#join20=pd.merge(join19,Nons_reasons,on="Orders_moved_to_NONS_Status_in_USS_due_to_COMS_status",how="left")
#join21=pd.merge(join20,hold_reasons,on="Hold_orders_issue_in_USS",how="left")
#join22=pd.merge(join21,IOS_errors,on="IOS_errored_out_orders",how="left")

final1=join19.drop_duplicates('Order_Number',keep='first')
final1.rename(columns={'LOC_CODE_x':'LOC_CODE'},inplace=True)
final1.drop('LOC_CODE_y',axis=1,inplace=True)
reqdata=final1[['Order_Number','LOC_CODE','Receive_Order','Master_ship_number','Routing_Order','PickList_Printing','PackList_Printing','Shipping_Label_Printing','ASN_label_printing','Allocation_Received','Planning_Releasing','BOL_Printing','Shipment_Date','ASN_Orders','inventory','Original_Exp_Ship_Date','Revised_Ship_date','REV_SHIP_DT_RSN_CD','REV_SHIP_DT_RSN','CUST_SHIP_TO_NBR','CUST_NAME','COUNTRY_CODE','Order_Class_Code','MMM_ID_NBR','USS_GOOD_SVC_TYPE_CODE','USS_Order_locked_timestamp','USS_Order_released_timestamp','Orders_moved_to_NONS_Status_in_USS_due_to_COMS_status','Reason','BRIEF_DESC','Hold_orders_issue_in_USS','Hold reason Description','IOS_errored_out_orders','IOS reason Description']]
reqdata.to_excel(path+"\\"+"BVM_Report.xlsx",index=False)
print("success")

