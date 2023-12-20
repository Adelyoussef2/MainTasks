import tkinter as tk
from tkinter import *
from PIL import ImageTk, Image
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import numpy as np


window = tk.Tk()
window.geometry("500x300")
window.title("Modification")
background_img = Image.open("oppo-reno-logo-a (2).jpg")
icon = PhotoImage(file="Oppo.jpg")


window.iconphoto(False, icon)


label = Label(window, text="Hello, world!")
label.pack()


image = ImageTk.PhotoImage(background_img)


background_label = Label(window, image=image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

file_path_1 = ''
file_path_2 = ''


def browse_file_1():
    global file_path_1
    file_path_1 = filedialog.askopenfilename()


def browse_file_2():
    global file_path_2
    file_path_2 = filedialog.askopenfilename()


def perform_analysis():
  
    data_1 = pd.read_excel(file_path_1)
    data_2 = pd.read_excel(file_path_2)
    df = pd.concat([data_1, data_2], ignore_index=True)
    s1=df[['Repair Order No.','SC Name','Repair type','Fault Phenomenon Code','Fault Cause Code']]
    s1=s1.loc[(s1['Repair type'].isin(['IW','IW&OOW'])) & 
       (s1['Fault Cause Code'].isin
        (['YQT7', 'PZJH6', 'ZBH4', 'RW5', 'JGJH16', 'JGJH3', 'PZJH3', 'PZJH2', 'PZJH1', 'JY2', 
          'JGJH1', 'YSQH2', 'JY3', 'JY5', 'FBH7', 'JY1', 'JY4', 'JY6', 'RW3', 'JGJH4', 'JGJH2']))]  
    s2=df.loc[(df['Repair type'].isin(['OOW'])) & (df['Net charge'] != df['Receivable price'])]
    s2=s2.dropna(subset=['Net charge'])
    s2=s2[['Repair Order No.','Service Content','Discount Content','Special Record No.']]  
    s3=df.fillna({'Good Material Subcategory':'Blank'})
    s3['Good Material Subcategory']
    s3=s3.loc[(s3['Good Material Subcategory']
        !='Protective Film') & (s3['Good Material Subcategory']
                                                        !='Auxiliary Material') & (s3['Good Material Subcategory']
                                                                                   !='Blank')]
    s3=s3.loc[s3['Solution']!= 'Material Replacement']
    s3=s3[['Repair Order No.','SC Code','Solution']]
    s4=df[['Repair Order No.','Good Material Subcategory','Repaired Time','Send to Repair Time',]]
    s4['Repaired Time']=pd.to_datetime(s4['Repaired Time'])
    s4['Send to Repair Time']=pd.to_datetime(s4['Send to Repair Time'])
    s4['Min']=s4['Repaired Time']-s4['Send to Repair Time']
    s4['Min'] = s4['Min']/np.timedelta64(1,'m')
    s4=s4.loc[s4['Good Material Subcategory'].isin(['Screen','Mainboard','Mainboard (8G 128G)','Mainboard (6G 128G)'])]
    s4=s4.loc[s4['Min']<10]
    s5=df[['Repair Order No.','Repair type','Receive Date','Purchase Date','Solution','Quality Record Category No.','Remark']]
    s5['Receive Date']=pd.to_datetime(s5['Receive Date'])
    s5['Purchase Date']=pd.to_datetime(s5['Purchase Date'])
    s5['Days']= (s5['Receive Date'] - s5['Purchase Date'])/np.timedelta64(1,"D")
    
    s5=s5.loc[(s5['Repair type'].isin(['IW','IW&OOW'])) & (s5['Days'] > 427) & (s5['Solution'] != 'Clean')] 
    s6=df[['Repair Order No.','Have defective materials been taken away or not']]
    s6=s6.loc[s6['Have defective materials been taken away or not'] == 'Yes']
    s7=df[['Repair Order No.','Repair type','Submission Category','Net charge','Receivable price']]
    s7=s7.loc[(s7['Repair type'].isin(['OOW'])) 
       & (s7['Submission Category'].isin(['Salesman','Dealer'])) 
       & (s7['Net charge'] != s7['Receivable price'])]
    s7=s7.fillna(0)
    s8=df[['Repair Order No.','Repair type','Submission Category','Receivable price','Net charge',]]
    s8=s8.loc[(s8['Repair type'].isin(['OOW'])) 
       & (s8['Submission Category'].isin(['Customer'])) 
       & (s8['Net charge'] == s8['Receivable price'])]
    s8['Discount']=s8['Net charge'] - s8['Receivable price']
    s8=s8.loc[s8['Discount'] != 0.0]
    s9= df[['Order Status','Handle Method','Customer Fault Description','Repair Order No.','Fault Phenomenon Code','Fault Cause']]
    s9=s9.loc[s9['Customer Fault Description']
       .str.contains('upgrade') | s9['Customer Fault Description']
       .str.contains('software') | s9['Customer Fault Description']
       .str.contains('سوفت') |s9['Customer Fault Description']
       .str.contains('تحديث')|s9['Customer Fault Description']
       .str.contains('soft ware') |s9['Customer Fault Description']
       .str.contains('سوفت وير') |s9['Customer Fault Description']
       .str.contains('سوفتوير')]
    s9=s9.loc[(s9['Handle Method']!= 'Cancel Service')&(s9['Order Status']!='Repairing')&(s9['Handle Method']!='Consultation')]
    s10=df[['Order Status','Handle Method','Customer Fault Description','Repair Order No.','Fault Phenomenon Code','Fault Cause']]
    s10=s10.loc[s10['Customer Fault Description'].str.contains('password')|
        s10['Customer Fault Description'].str.contains('باسوورد') |
        s10['Customer Fault Description'].str.contains('باسورد') |
        s10['Customer Fault Description'].str.contains('مرور')|
        s10['Customer Fault Description'].str.contains('السر') |
        s10['Customer Fault Description'].str.contains('lockscreen') ]
    s10=s10.loc[(s10['Handle Method']!= 'Cancel Service')&(s10['Order Status']!='Repairing')&(s10['Handle Method']!='Consultation')]
    s11=df[['Repair Order No.','Repair type','Order Type','Receive Date','Purchase Date','Solution']]
    s11['Receive Date']=pd.to_datetime(s11['Receive Date'])
    s11['Purchase Date']=pd.to_datetime(s11['Purchase Date'])
    s11['Days']= (s11['Receive Date'] - s11['Purchase Date'])/np.timedelta64(1,"D")
    
    s11=s11[(s11['Order Type']=='IoT')&((s11['Receive Date']-s11['Purchase Date'])>'365 days') & (s11['Repair type'].str.contains('IW')==True)]
    s12=df[(df['Marketing Model']=='OPPO Reno6')&(df['Repair type'].str.contains('OOW|IW&OOW|IW')==True)&(df['Good Material Code']==4906118)]
    s12=s12[['Repair Order No.','SC Name','Marketing Model','Good Material Code']]
    s13=df[(df['Marketing Model']=='OPPO Reno6')&(df['Repair type'].str.contains('IW')==True)&(df['Good Material Code']==4906118)&(df['Remark'].isnull()==True)]
    s13=s13[['Repair Order No.','SC Name','Marketing Model','Good Material Code','Remark']]
    s15=df[(df['Good Material Subcategory'].str.contains('screen|Screen|Screen|Mainboard'))&(df['QR code for Good Material'].isnull()==True)&(df['Remark for Replace Materials'].isnull()==True)&(df['Handle Method']!='Cancel Service')&(df['Order Status']!='Repairing')&(df['Good Material Code'].isnull()==False)]
    s15=s15[['Repair Order No.','SC Name','Order Status','Handle Method','QR code for Good Material','Good Material Subcategory','Remark for Replace Materials']]
    s16=df[df['Special Record Status']=='Not submit']
    s17=df[(df['Repair type']=='OOW')&((df['Receivable price']-df['Net charge'])!=0)&(df['Discount Content']!='Phone Protection Plan')]
    s17['percentile of discount']=(s17['Net charge']-s17['Receivable price'])/s17['Receivable price']
    s17=s17[['Repair Order No.','SC Name','Discount Content','Special Record No.','Service Content','percentile of discount']]
    
    

    
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=(("Excel files", "*.xlsx"),))

    
    with pd.ExcelWriter(save_path) as writer:
        sheet_name = "IW"
        s1.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Net Change"
        s2.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "Solution"
        s3.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "10 MIN"
        s4.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "427 Day"
        s5.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "Defect Taken"
        s6.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "discount for dealer"
        s7.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "discount"
        s8.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "SW1"
        s9.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "SW3"
        s10.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "IoT exceed"
        s11.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "reno screen"
        s12.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "Reno remark"
        s13.to_excel(writer, sheet_name=sheet_name, index=False)

        sheet_name = "QR"
        s15.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "Special issue"
        s16.to_excel(writer, sheet_name=sheet_name, index=False)
        
        sheet_name = "balance"
        s17.to_excel(writer, sheet_name=sheet_name, index=False)






button_1 = tk.Button(window, text="Handover file", command=browse_file_1)
button_2 = tk.Button(window, text="Receive File", command=browse_file_2)
button_3 = tk.Button(window, text="Start", command=perform_analysis)


button_1.pack()
button_1.place(x=210, y=30)
button_2.pack()
button_2.place(x=215, y=100)
button_3.pack()
button_3.place(x=235, y=180)


window.mainloop()
