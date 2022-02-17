import os
os.environ["KMP_DUPLICATE_LIB_OK"]="TRUE"
import rpa as r   
import pyautogui 
import easyocr
import pandas as pd 
import subprocess
import numpy as np
from datetime import datetime

reader = easyocr.Reader(['en'],gpu=False)
#r.init(visual_automation = True, chrome_browser = False)
        
summary = {"Invoice No":[],"Description":[],"Quantity":[],"Unit_Price":[],
           "Amount":[]}


#import path locations
from path_find import main_file
mainfile = main_file

ref_path = pd.read_excel(mainfile, sheet_name="References")\
    .drop(columns=['Description'])\
        .set_index("Object")["Identifier"]\
            .to_dict()
         
#datetime now
now = datetime.now()
date_time = now.strftime("%d%m%y-%H%M")
# anchor images location 
img_loc = str(ref_path["anchor_loc"]+"\\Chinyuan")

#invoice location
invoice_loc = ref_path["i_path"]

#importing masterfile with company name 
df = pd.read_excel("output.xlsx", sheet_name="Sheet1")
      

def description ():
    
    for box in pyautogui.locateAllOnScreen(img_loc+'\\zbnw.png'):
        
        left,top,width,height = box
      
        description = pyautogui.screenshot(region=(left+3,top+27,width+690,height+28))
        description.save('out1.png','PNG')
        bound1= str(reader.readtext('out1.png',detail=0))
        
        summary["Invoice No"].append(bound0)        
        summary["Description"].append(bound1)

def quantity ():
    
    for box in pyautogui.locateAllOnScreen(img_loc+'/zbnw.png'):
        
        left,top,width,height = box
        
        quantities = pyautogui.screenshot(region=(left+735,top,width+45,height))
        quantities.save('out2.png','PNG')
        bound2= str(reader.readtext('out2.png',detail=0))
        summary["Quantity"].append(bound2)
        
def unit_price ():
    for box in pyautogui.locateAllOnScreen(img_loc+'/zbnw.png'):
        
        left,top,width,height = box
        
        unitpr = pyautogui.screenshot(region=(left+915,top,width+39,height))
        unitpr.save('out3.png','PNG')
        bound3= str(reader.readtext('out3.png',detail=0))
        summary["Unit_Price"].append(bound3)

def amount ():
    for box in pyautogui.locateAllOnScreen(img_loc+'/zbnw.png'):
        
        left,top,width,height = box
        
        unitpr = pyautogui.screenshot(region=(left+1053,top,width+58,height+5))
        unitpr.save('out4.png','PNG')
        bound4= str(reader.readtext('out4.png',detail=0))
        summary["Amount"].append(bound4)


for i in range (0,df.shape[0],1):
    
    
    if df.iloc[i][3] == 'Chinyuan':
        count = 0
        subprocess.Popen([ref_path["i_path"]+"\\"+df.iloc[i][1]],shell=True) 
        r.wait(2)
        r.keyboard("[home]")
        
        #Extracting Invoice No 
        left,top,width,height = pyautogui.locateOnScreen(img_loc +'\\invoice.png')
        invoiceno = pyautogui.screenshot(region=(left+213,top,width+45,height+3))
        invoiceno.save('out0.png','PNG')
        bound0= str(reader.readtext('out0.png',detail=0))  
        
        while count<10:
            
                description()
                quantity()
                unit_price()
                amount()
                r.wait(1)
                
                if pyautogui.locateOnScreen(img_loc+'\\total.png',confidence=.8):
                    break
                elif pyautogui.locateOnScreen(img_loc+'\\total.png',confidence = .8) is None:
                    r.keyboard('[pagedown]')
                    count += 1
                    
        else: break
    
df1 = pd.DataFrame(summary)
df1 = df1.applymap(lambda x : x[2:-2])
df1['Quantity'] = df1['Quantity'].str[:-3]
df1["Quantity"] = df1.Quantity.astype(float)
df1["Unit_Price"] = df1.Unit_Price.astype(float)
df1["Amount"] = df1.Amount.astype(float)
df1['Amount'] = df1.Amount.round(2)

df1['Checked_amount'] = df1.apply(lambda row: row.Quantity * 
                                  row.Unit_Price, axis = 1)
df1['Checked_amount'] = df1.Checked_amount.round(2)
df1['Test_Result'] = np.where(df1['Amount']==df1['Checked_amount'], True, False)
df1.drop("Checked_amount",axis=1,inplace=True)
df1['Description']=df1['Description'].str.replace("'","")
sn =df1['Description'].str.split(")", n = 1, expand = True)
df1['SN'] = sn[0]
df1['Description']=sn[1]
df1 = df1[['Invoice No','SN','Description','Quantity','Unit_Price','Amount','Test_Result']]

writer = pd.ExcelWriter('chinyuan.xlsx')
df1.to_excel(writer)
writer.save()
df1.to_excel(ref_path["p_path"]+"\\ScanOutput\\Chinyuan\\Chinyuan_"+date_time+".xlsx")

"""
add to history
"""
from openpyxl import load_workbook
name = ("Chinyuan_"+date_time,)
wb = load_workbook('History.xlsx')
page = wb.active
page.append(name)
wb.save(filename='History.xlsx')



