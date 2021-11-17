# Meastro to Insight
# Written by: Anthony Bradt 613-986-0029
# Requested by: Victoria Hurrell


import pyodbc
import pandas as pd

sqlServer = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=LK-s-erp;"
                      "Database=inSight;"
                      "Trusted_Connection=yes;")                                            

#Put the Fun in Functions
def SQLRead():
    try:#SQL provided by Tommy
        SQLWO = "SELECT ordPONumber FROM dbo.Orders WHERE ordPONumber LIKE '[1-9]%' AND ordPONumber NOT LIKE '%-%'"
        workOrdersDF = pd.read_sql(SQLWO,sqlServer)
        return workOrdersDF
    except:#If Failed, Try again. Bandain for timeout SQL Requests
        print("SQL failed")
        SQLRead()

def MaestroRead():
    maestroDF = pd.read_excel(r'C:\Users\AnthonyB\Desktop\scripts\maestro.xls')
    #print (meastroDF['Customer'].unique())\
    #meastroDF =  meastroDF.drop(['4 - Final app.'])
    #maestroDF = maestroDF[['Work Order','Status','Customer']]
    maestroDF = maestroDF.loc[maestroDF['Status'] == '2 - Confirmed',['Work Order','Status','Customer','Acctg Date','Description - Employee in Charge','Description','Lot # (BTRAV.LOT01703)','Site (BTRAV.SIT0293)']]

    return maestroDF
   
#Main
input("Step 1 -- Press enter to continue")
inSightData = SQLRead()
#inSightData = inSightData.astype('int64')
print (inSightData)
maestroData = MaestroRead()
#maestroData = maestroData.astype('object')
for index, row in maestroData.iterrows():
    WO = row['Work Order']
    #CU = row['Customer']
    if str(WO) in inSightData.values:
       #print ("Already Exists")
        x = 1
    else:
        if 'UPGRADES' in row['Description'].upper():
            x = 1
            #print("We dont add upgrades to insight")
        elif 'LABOUR' in row['Description'].upper():
            x = 1
            #print("We dont add labour to insight")
        elif 'TOUCH UP' in row['Description'].upper():
            x = 1
        elif 'TOUCH' in row['Description'].upper():
            x = 1
        else:
            print(row['Work Order'], row['Customer'], row['Acctg Date'], row['Description - Employee in Charge'], row['Description'],row['Lot # (BTRAV.LOT01703)'],row['Site (BTRAV.SIT0293)'])
print (maestroData)



input("Finished")
