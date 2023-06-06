# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#import os
import win32com.client as wc
# import xlsxwriter as xls
import pandas as pd
import time
from enum import Enum


class sensitivity_enum(Enum):
    Normal = 0
    Personal = 1
    Private = 2
    Confidential = 3
    
    
class importance_enum(Enum):
    Low = 0
    Normal = 1
    High = 2
    
#--Starting outlook session--
#connecting to Outlook by MAPI
outlook = wc.Dispatch("Outlook.Application").GetNamespace("MAPI")
print("Starting outlook session")
   
def Mask_Email(email):
    #finding the location of @
    loc = email.find('@')
    if loc>0:
        email = email[0]+"********"+email[loc-1:]
        return email
    else:
        return "Invalid Email"
        

def export_inbox_data():
    
    print("Inside export_inbox_data() ")
    #accessing inbox folder
    inbox = outlook.GetDefaultFolder(6)  
    #accessing inbox folder messages
    messages = inbox.Items
    # "Subject",   "Body",   "From (Name)",  "From (Address)",      "From (Type)",     "Receiver",      "CC",      "Bcc",  "Category",  "Sensitivity",      "Importance"
    #Subject    Body    From: (Name)    From: (Address)    From: (Type)    To: (Name)    To: (Address)    To: (Type)    CC: (Name)    CC: (Address)    CC: (Type)    BCC: (Name)    BCC: (Address)    BCC: (Type)    Billing Information    Categories    Importance    Mileage    Sensitivity                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            

    Subject = []
    Body = []
    Received_Time = []
    From_Name = []
    From_Address = []
    From_Type= []
    To_Name = []  
    To_Address = []
    To_Type= []
    Cc_Name = []
    Cc_Address =[]
    Cc_Type=[]
    Bcc_Name =[]
    Bcc_Address =[]
    Bcc_Type=[]
    Category =[]
    Importance=[]
    Mileage= []
    Sensitivity=[]
    
    num_items=0
    email = None
    
    
    
    # starting time
    st = time.time()
    
    for msg in list(messages)[0:20]:
        
       #i=0
        #eliminating meeting invites
        if msg.Class==43:
            
            num_items= num_items + 1 
            print("Item", num_items)
            #checking Exchange User Type sender
            if msg.SenderEmailType=='EX':
                
                if msg.Sender.GetExchangeUser() != None:
                    email = msg.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    email = msg.Sender
                    
            else:
                # SMTP user type
                email = msg.SenderEmailAddress
             

            recip_name=[]
            recip_address=[]
            recip_address_type=[]
   
            cc_recip_name=[]
            cc_recip_address=[]
            cc_recip_address_type=[]
            
            bcc_recip_name=[]
            bcc_recip_address=[]
            bcc_recip_address_type=[]
              
            # for recipients
            recipients = msg.Recipients
            
            
            for recipient in recipients:                                
                # recipient.Type 1-to; 2-cc; 3-bcc
                
                if recipient.Type == 1:
                    try:                        
                        recipient_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    except AttributeError:
                        recipient_address = recipient.AddressEntry.Address                    
                                                     
                    recip_name.append(str(recipient))
                    recip_address.append(recipient_address)
                    recip_address_type.append(recipient.AddressEntry.Type)                    
                
                elif recipient.Type == 2:                        
                    try:
                        cc_recipient_address= recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress                   
                    except AttributeError:
                        cc_recipient_address = recipient.AddressEntry.Address
                    
                         
                    cc_recip_name.append(str(recipient))
                    cc_recip_address.append(cc_recipient_address)
                    cc_recip_address_type.append(recipient.AddressEntry.Type)                            
                        
                elif recipient.Type == 3:                      
                    try:
                        bcc_recipient_address= recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress                         
                    except AttributeError:
                        bcc_recipient_address = recipient.AddressEntry.Address
                    
                    
                    bcc_recip_name.append(str(recipient))
                    bcc_recip_address.append(bcc_recipient_address)
                    bcc_recip_address_type.append(recipient.AddressEntry.Type)
                    
                    
                else:
                    print("Something is wrong")
                    
            print("recip_name",recip_name)     
                    
            Subject.append(msg.Subject)
            
            # an excel cell accepts only maximum 32767 chars
            Body.append(msg.Body[:32766])
            
            try:
                Received_Time.append(str(msg.ReceivedTime))
            except AttributeError:
                Received_Time.append("NA")
            
            
            try:            
                From_Name.append(str(msg.Sender))
            except AttributeError:
                From_Name.append("NA")
             
            From_Address.append(email)
    
            #476
            try:            
                From_Type.append(msg.SenderEmailType)
            except AttributeError:
                From_Type.append("NA")
            
            
            
            To_Name.append(recip_name)            
            To_Address.append(recip_address)
            To_Type.append(recip_address_type)
            print("To Name",To_Name)           
           
                       
            Cc_Name.append(cc_recip_name)
            Cc_Address.append(cc_recip_address)
            Cc_Type.append(cc_recip_address_type)
            
            
                        
            Bcc_Name.append(bcc_recip_name)
            Bcc_Address.append(bcc_recip_address)
            Bcc_Type.append(bcc_recip_address_type)                       
            
            
            Category.append(msg.Categories)
            
            b = msg.Importance
            Importance.append(importance_enum(b).name)
            
            Mileage.append(msg.Mileage)
            
            a = msg.Sensitivity
            Sensitivity.append(sensitivity_enum(a).name)
            
            # since to, cc, bcc are nested lists, thus, clearing the inner lists
            recip_name.clear()
            recip_address.clear()
            recip_address_type.clear()
            
            cc_recip_name.clear()
            cc_recip_address.clear()
            cc_recip_address_type.clear()
            
            bcc_recip_name.clear()
            bcc_recip_address.clear()
            bcc_recip_address_type.clear()
        
# =============================================================================
#         i = i+1
# =============================================================================
    et = time.time()       
    elapsed_time = et - st
    print('Execution time:', elapsed_time, 'seconds')
    print("total number of inbox items: ",  num_items)
    print("Subject count", len(Subject))
    print("To_Name count", len(To_Name))
    print("To_Address count", len(To_Address))
    print("To_Type count", len(To_Type))
    
    #creating dataframe
    #Subject    Body    From: (Name)    From: (Address)    From: (Type)    To: (Name)    To: (Address)    To: (Type)    CC: (Name)    CC: (Address)    CC: (Type)    BCC: (Name)    BCC: (Address)    BCC: (Type)    Billing Information    Categories    Importance    Mileage    Sensitivity                                                   
    df = pd.DataFrame()
    df['Subject'] = Subject
    df['Body'] = Body
    df['Received Time'] = Received_Time
    df['From: (Name)'] = From_Name
    df['From: (Address)'] = From_Address
    df['From: (Type)'] = From_Type
    
    df['To: (Name)'] = To_Name
    
    df['To: (Address)'] = To_Address
    
    df['To: (Type)'] = To_Type
    df['CC: (Name)'] = Cc_Name
    df['CC: (Address)'] = Cc_Address
    df['CC: (Type)'] = Cc_Type
    df['BCC: (Name)'] = Bcc_Name
    df['BCC: (Address)'] = Bcc_Address
    df['BCC: (Type)'] = Bcc_Type
    df['Categories'] = Category
    df['Importance'] = Importance
    df['Mileage'] = Mileage
    df['Sensitivity'] = Sensitivity
    
    print("Export Started")
    # Creating Excel file with results
    #cwd = os.getcwd()
    #df.to_csv(r'{0}\inbox_dataset.csv'.format(cwd))
    df.to_excel(r'C:\Users\ismgupta\Desktop\lol\inbox.xlsx')
    print("Export Successfully completed")
    
    
#------------------------------------------------------------#

def export_calendar_data():
    calendar = outlook.GetDefaultFolder(9) 
    events= calendar.Items
    
    Subject = []
    All_day_event = []
    Organizer = []
    Organizer_Email_Address = []
    Required_Attendees = []
    Optional_Attendees = []
    Categories = []
    Description = []
    Location = [] 
    Mileage = []
    Sensitivity = []
    Creation_Time = []
    
    num_items =0
    
    st = time.time()
    for event in events:
        c=event.GetOrganizer()
        num_items = num_items +1
        
        Subject.append(event.Subject)
        All_day_event.append( event.AllDayEvent)
        Organizer.append(event.Organizer)
        Organizer_Email_Address.append(str(c.GetExchangeUser().PrimarySmtpAddress))
        Required_Attendees.append(event.RequiredAttendees)
        Optional_Attendees.append(event.OptionalAttendees)		
        Categories.append(event.Categories)
        Description.append(event.Body[:32766])
        Location.append(event.Location) 
        Mileage.append(event.Mileage)
        Sensitivity.append(event.Sensitivity)
        Creation_Time.append(str(event.CreationTime))
       
    et = time.time()       
    elapsed_time = et - st
    print('Execution time:', elapsed_time, 'seconds')
    print("total number of inbox items: ",  num_items)
    
    print("Creating Data Frame")
    df = pd.DataFrame()
    df['Subject'] = Subject
    df['Creation Time'] = Creation_Time
    df['Meeting Organizer'] = Organizer
    df['Meeting Organizer : Email'] = Organizer_Email_Address
    df['Required Attendees'] = Required_Attendees
    df['Optional Attendees'] = Optional_Attendees
    df['Categories'] = Categories
    df['Description'] = Description
    df['Location'] = Location
    df['Mileage'] = Mileage
    df['Sensitivity'] = Sensitivity
    df['All day event'] = All_day_event
    
    
# =============================================================================
#     print("Configure Columns")
#     print(df.columns)
#     
#     drop_columns = []
#     
#     drop_columns.append(input("Enter the names of the column that you want to drop"))
#    
#     df.drop(drop_columns)
# =============================================================================
    
    #for column in df.columns:
        #if drop_columns.Contains(column):
            #df.drop[column]
    
    print("Export Started")
    # Creating Excel file with results
    #cwd = os.getcwd()
    #df.to_csv(r'{0}\inbox_dataset.csv'.format(cwd))
    df.to_excel(r'C:\Users\ismgupta\Desktop\lol\calendar.xlsx')
    print("Export Successfully completed")
    
def export_sent_items():
    sentItems = outlook.GetDefaultFolder(5)
    messages = sentItems.Items
    
    num_items =0
    st = time.time()
    # Subject	Body	From: (Name)	From: (Address)	From: (Type)	To: (Name)	To: (Address)	To: (Type)	CC: (Name)	CC: (Address)	CC: (Type)	BCC: (Name)	BCC: (Address)	BCC: (Type)	Billing Information	Categories	Importance	Mileage	Sensitivity
    
    Subject = []
    Body = []
    Received_Time = []
    From_Name = []
    From_Address = []
    From_Type= []
    To_Name = []  
    To_Address = []
    To_Type= []
    Cc_Name = []
    Cc_Address =[]
    Cc_Type=[]
    Bcc_Name =[]
    Bcc_Address =[]
    Bcc_Type=[]
    Category =[]
    Importance=[]
    Mileage= []
    Sensitivity=[]
        
    for msg in messages:
        num_items = num_items +1
         
    if msg.Class==43:
        
        num_items= num_items + 1 
        print("Item", num_items)
        #checking Exchange User Type sender
        if msg.SenderEmailType=='EX':
            
            if msg.Sender.GetExchangeUser() != None:
                email = msg.Sender.GetExchangeUser().PrimarySmtpAddress
            else:
                email = msg.Sender
                
        else:
            # SMTP user type
            email = msg.SenderEmailAddress
         

        recip_name=[]
        recip_address=[]
        recip_address_type=[]

        cc_recip_name=[]
        cc_recip_address=[]
        cc_recip_address_type=[]
        
        bcc_recip_name=[]
        bcc_recip_address=[]
        bcc_recip_address_type=[]
          
        # for recipients
        recipients = msg.Recipients
        
        
        for recipient in recipients:                                
            # recipient.Type 1-to; 2-cc; 3-bcc
            
            if recipient.Type == 1:
                try:                        
                    recipient_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                except AttributeError:
                    recipient_address = recipient.AddressEntry.Address                    
                                                 
                recip_name.append(str(recipient))
                recip_address.append(recipient_address)
                recip_address_type.append(recipient.AddressEntry.Type)                    
            
            elif recipient.Type == 2:                        
                try:
                    cc_recipient_address= recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress                   
                except AttributeError:
                    cc_recipient_address = recipient.AddressEntry.Address
                
                     
                cc_recip_name.append(str(recipient))
                cc_recip_address.append(cc_recipient_address)
                cc_recip_address_type.append(recipient.AddressEntry.Type)                            
                    
            elif recipient.Type == 3:                      
                try:
                    bcc_recipient_address= recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress                         
                except AttributeError:
                    bcc_recipient_address = recipient.AddressEntry.Address
                
                
                bcc_recip_name.append(str(recipient))
                bcc_recip_address.append(bcc_recipient_address)
                bcc_recip_address_type.append(recipient.AddressEntry.Type)
                
                
            else:
                print("Something is wrong")
                
        print(recip_name)     
                
        Subject.append(msg.Subject)
        
        # an excel cell accepts only maximum 32767 chars
        Body.append(msg.Body[:32766])
        
        try:
            Received_Time.append(str(msg.ReceivedTime))
        except AttributeError:
            Received_Time.append("NA")
        
        
        try:            
            From_Name.append(str(msg.Sender))
        except AttributeError:
            From_Name.append("NA")
         
        From_Address.append(email)

        #476
        try:            
            From_Type.append(msg.SenderEmailType)
        except AttributeError:
            From_Type.append("NA")
        
        
        
        To_Name.append([recip_name])            
        To_Address.append([recip_address])
        To_Type.append([recip_address_type])
                    
       
                   
        Cc_Name.append(cc_recip_name)
        Cc_Address.append(cc_recip_address)
        Cc_Type.append(cc_recip_address_type)
        
        
                    
        Bcc_Name.append(bcc_recip_name)
        Bcc_Address.append(bcc_recip_address)
        Bcc_Type.append(bcc_recip_address_type)                       
        
        
        Category.append(msg.Categories)
        
        b = msg.Importance
        Importance.append(importance_enum(b).name)
        
        Mileage.append(msg.Mileage)
        
        a = msg.Sensitivity
        Sensitivity.append(sensitivity_enum(a).name)
        
        # since to, cc, bcc are nested lists, thus, clearing the inner lists
        recip_name.clear()
        recip_address.clear()
        recip_address_type.clear()
        
        cc_recip_name.clear()
        cc_recip_address.clear()
        cc_recip_address_type.clear()
        
        bcc_recip_name.clear()
        bcc_recip_address.clear()
        bcc_recip_address_type.clear()
    
    et = time.time()       
    elapsed_time = et - st
    print('Execution time:', elapsed_time, 'seconds')
    print("total number of inbox items: ",  num_items)
    
   
#def Select_Task():
    #while True:

        #class Python_Switch:
            #def day(self, response):
        
                #default = "Done with the selection"
        
                #return getattr(self, 'case_' , lambda: default)()
        
            #def case_1(self):
                #export_inbox_data()
        
            #def case_2(self):
                #export_calendar_data()
        
            #def case_3(self):
                #export_senditems_data()
            
        #def case_4(self):
                #export_all_items()
                
    #def case_5(self):
                #exit()
    
    
    #my_switch = Python_Switch()
    
    #print(my_switch.day(1))   
    

   
def main():
   #retrieving all outlook accounts in your system
   accounts= wc.Dispatch("Outlook.Application").Session.Accounts;
   for account in accounts:
       print(account.DeliveryStore.DisplayName) 
       export_inbox_data()    
       export_calendar_data()
       print("closing the accounts")
       for account in accounts:
            print(account.DeliveryStore.DisplayName)
            account.Application.Quit()
            print("account closed")

# =============================================================================
#    print("enter 0 for exporting inbox emails")
#    print("enter 1 for exporting calendar invites")
#    print("enter 2 for exporting sent items")
#    print("enter 3 for exporting all of the above in one workbook")
#    print("enter 9 for exit")
#            
#    response = input("please enter your response")
#            
#    if response == 0:
#        export_inbox_data()
#        
#    else:
#        print("enter 0")
# =============================================================================
   
  

if __name__ == "__main__":
    main()        
    


    

