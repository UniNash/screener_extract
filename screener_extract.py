

#-------------importing libraries------------

import requests
import os
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime as dt
import sys
import time
import openpyxl
import shutil


start = time.time()  # timer to check execution time

# function to define countdown timer
-------------------------------------

# time_sec - seconds you want countdown for.
def countdown(time_sec):
    while time_sec: # no need to specify condition; when time_sec reaches 0 - becomes False; 0 represents False
        mins, secs = divmod(time_sec, 60)   # gives the remainder(min) and quotient(secs)
        timeformat = '{:02d}:{:02d}'.format(mins, secs) # formatting 
        print(timeformat, end='\r') # move the cursor or print head back to the beginning of the current line and print time
        time.sleep(1) # 1 second delay is used to avoid instant printing.
        time_sec -= 1
    print("done")


# function to login screener.com 
-------------------------------------
def screener_login(url,loginextn,userid,user_password):
    s=requests.Session()
    main_url = url
    login_sup = loginextn
    home_url = os.path.join(main_url,login_sup)
    resp=s.get(home_url)
    if 'csrftoken' in resp.cookies:
        # Django 1.6 and up
        csrftoken = resp.cookies['csrftoken'] # uses csrftoken  CSRF (Cross Site Request Forgery) 
    else:
        # older versions
        csrftoken = resp.cookies['csrf']
    login_data = dict(username=userid, password=user_password, csrfmiddlewaretoken=csrftoken, next='/') # creating a dictionary of login keys and inputs
    r = s.post(home_url, data=login_data, headers=dict(Referer=home_url)) # returns 
    print(s)
    return s  #return the sessions object


#required ratios from screener
#mapping dictionary label in website and column name in file

required_ratios = {'Market Cap':'M Cap (in Cr.)'
                   ,'Current Price':'Sh Pr'
                   ,'Stock P/E':'PE'
                   ,'Price to book value':'P/B'
                   ,'ROCE':'ROCE'
                   ,'ROE':'ROE'
                   ,'Return on assets':'ROA'
                   ,'Debt to equity':'DTE'
                   ,'OPM':'OPM'
                   ,'Earnings yield':'EY'
                   ,'PEG Ratio':'PEG'
                   ,'Dividend Yield':'DY'
                   ,'10 Years:':'10 Yrs'
                   ,'5 Years:':'5 Yrs'
                   ,'3 Years:':'3 Yrs'
                   ,'TTM:':'TTM'}

columns_list = ['Trigger','M Cap (in Cr.)','Sh Pr','PE','P/B','ROCE', 'ROE','ROA', 'DTE','OPM', 'EY','PEG', 'DY','10 Yrs','5 Yrs','3 Yrs','TTM',10.0,5.0,3.0,'TTMc']

main_url = 'https://www.screener.in/'
login_sup = 'login/?'

print("------------------------------------------------------------------------------")
print('\n')
print("Enter your file path:")                 
file_path = input()                     # file path of the input/output file
print("------------------------------------------------------------------------------")
print("File Path: "+ file_path)
print("------------------------------------------------------------------------------")

# taking a backup of input file, because the original file will be overwritten

to_path = file_path[0:len(file_path)-5]+"_"+"backed_up_on"+str(dt.today().date())+".xlsx"
shutil.copy(file_path, to_path) # copy the file from source to a destination directory (which is the same in this case)
print("------------------------------------------------------------------------------")
print("File backup successful.",to_path)
print("------------------------------------------------------------------------------")


try:
    actual_file = pd.read_excel(file_path) #reading the excel file
    print("------------------------------------------------------------------------------")
    print("File successfully uploaded.")
    print("------------------------------------------------------------------------------")
except FileNotFoundError:
    print("------------------------------------------------------------------------------")
    print("File unavailable.Please place the file.")
    print("------------------------------------------------------------------------------")
    sys.exit() #exiting the program
except ValueError:
    print("------------------------------------------------------------------------------")
    print("Unrecognized file format.Please provide the correct file.")
    print("------------------------------------------------------------------------------")
    sys.exit() #exiting the program

if(list(actual_file.loc[0][0:21])==columns_list): #checking the columns as per columns list
    print("------------------------------------------------------------------------------")
    print("No issues with file columns.Proceeding with data extraction")
    print("------------------------------------------------------------------------------")
else:
    print("------------------------------------------------------------------------------")
    print("File column names are incorrect, please upload the correct file and proceed")
    print("------------------------------------------------------------------------------")
    sys.exit() #exiting the program


print("Enter your Screener User ID: ")
user_id = input()
print("Enter your Screener Password: ")
pass_word = input()

print("------------------------------------------------------------------------------")
print("Userid: "+ user_id)
print("Pass: "+ pass_word)
print("------------------------------------------------------------------------------")

actual_file.columns=actual_file.loc[0] # get the first row of the table and make it as column headers

actual_file = actual_file.loc[:, actual_file.columns.notnull()] # checks and brings the values for non null columns

actual_file.index=[actual_file['Trigger']] # Trigger is nothing but the name of the companies to extract data from and is being changed to index.

actual_file.index.names = ['index'] # changing the name of the Index to "index"

#actual_file.drop(actual_file.index[:1],inplace=True)

actual_file = actual_file.drop(actual_file.index[:1]) # dropping the first row which is the same as the header name

print("------------------------------------------------------------")
print("File preprocessed successfully.")
print("------------------------------------------------------------")

search_list = list(actual_file["Trigger"]) # getting the trigger(company codes to search as list)

'''
for i in range (0,len(search_list)):
    try:
        if(pd.isna(search_list[i]+search_list[i+1]) == True):
            search_list_final = search_list[0:i]
            break
    except TypeError:
        pass
    except IndexError:
        pass
'''

# The trigger column contains NaN values and have to be removed

search_list_final = []
for i in range (0,len(search_list)):
    try:
        if(pd.isna(search_list[i]) == False):
            search_list_final.append(search_list[i])
    except TypeError:
        pass
    except IndexError:
        pass

# validation if the search list is empty

if(search_list_final==[]): 
    print("------------------------------------------------------------------------------")
    print('Search Company Empty in file. Please upload the file with data.')
    print("------------------------------------------------------------------------------")
    sys.exit()
else:
    search_list_final=[x for x in search_list_final if x==x]
    print('\n')
    print("------------------------------------------------------------------------------")
    print("{} Companies to extract".format(len(search_list_final)))
    print("------------------------------------------------------------------------------")

print("------------------------------------------------------------------------------")
s=screener_login(main_url,login_sup,user_id,pass_word)  # logging in to screener website
print("------------------------------------------------------------------------------")

ratio_value_list=[]
counter1=[]
counter2=[]
nodata=[]
ratio_value_1up=[]

# for every company code in the list the ratios or data is extracted from screener page or the company
# this block is to extract the basic ration from the home page

for company in search_list_final:
    if (len(counter1)==9):
        counter1=[]
        end = time.time()
        execution_time = (end - start)/60
        print("------------------------------------------------------------------------------")
        print("Time taken: {} mins".format(execution_time))
        print("Started..........")
        print("------------------------------------------------------------------------------")
    url_extn = 'company/'+ company  #+'/consolidated/'

    get_webpage=s.get(os.path.join(main_url,url_extn))
    print(os.path.join(main_url,url_extn))

    while(str(get_webpage)=='<Response [429]>'):       # since only limited no. of requests are allowed to Screener, we are creating a 20 sec time out for each request limit and then running the code.
        print("------------------------------------------------------------------------------")
        print("print-1")
        print("<Response [429]>': Cannot handle too may requests, wait for 20 seconds")
        print("------------------------------------------------------------------------------")
        #time.sleep(20)
        countdown(20)
        os.system('cls')
        print("Started..........")
        print("------------------------------------------------------------------------------")
        print(str(len(counter2))+" Companies extracted")
        print("------------------------------------------------------------------------------")
        get_webpage=s.get(os.path.join(main_url,url_extn))
    if(str(get_webpage)=='<Response [404]>'):         # if no data/page is found for the company no data is returned.
        print(company+" not found")
        nodata.append(company)
        continue
    counter1.append(company)
    counter2.append(company)
    parse_web_page_data = BeautifulSoup(get_webpage.content, 'html.parser')
    
    #default ratios
    default_ratio_section = parse_web_page_data.find(id="top-ratios") #get the default ratio data section
    default_ratio_items = default_ratio_section.select(".name") #get the default ratio names from webpage above (with html tags)
    default_ratio_values = default_ratio_section.select(".number") #get the default ratio values from webpage above (with html tags)

    # this block is to extract the quick ratios which are not in home page and have to be extracted using datawarehouseid api
    #------------------------------------------------------------------------------------------------------------------    
    #-----------------------------------------------quick ratios-------------------------------------------------------

    #---old code---
    '''
    datawarehouseid_tag=parse_web_page_data.main.div
    datawarehouseid = datawarehouseid_tag['data-warehouse-id'] #get warehouse id
    '''
    #--------------
    
    # new code

     # trying to get the datawarehouseid from the consolidated webpage data
     
    datawarehouseid  = parse_web_page_data.find_all("div", id="company-info")[0]["data-warehouse-id"]
    
    # if no datawarehouseid is found from the consolidated webpage data then using the normal web page url to extract the datawarehouse id
    if (datawarehouseid=='None'): 
        url_extn = 'company/'+company+'/'
        get_webpage=s.get(os.path.join(main_url,url_extn))
        if(str(get_webpage)=='<Response [404]>'):
            print("------------------------------------------------------------------------------")
            print(company+" warehouseid could not be found to extract quick ratios")
            print("------------------------------------------------------------------------------")
            nodata.append(company)
            continue
        while(str(get_webpage) =='<Response [429]>'):
            print("------------------------------------------------------------------------------") 
            print("print-2")
            print("<Response [429]>': Cannot handle too may requests, wait for 20 seconds")
            print("------------------------------------------------------------------------------")
            #time.sleep(20)
            countdown(20)
            os.system('cls')
            print("Started..........")
            print("------------------------------------------------------------------------------")
            print(str(len(counter2))+" Companies extracted")
            print("------------------------------------------------------------------------------")
            get_webpage=s.get(os.path.join(main_url,url_extn))

    
    parse_web_page_data = BeautifulSoup(get_webpage.content, 'html.parser')
        #---old code---
    '''
    datawarehouseid_tag=parse_web_page_data.main.div
    datawarehouseid = datawarehouseid_tag['data-warehouse-id'] #get warehouse id
    '''

    # 
    datawarehouseid  = parse_web_page_data.find_all("div", id="company-info")[0]["data-warehouse-id"]

    
    quick_ratio_url = os.path.join(main_url,'api/company/'+datawarehouseid+'/quick_ratios/')#create api url
    quick_ratio_page = s.get(quick_ratio_url)#get the api webpage
    
    while(str(quick_ratio_page) =='<Response [429]>'):
        print("------------------------------------------------------------------------------")
        print("print-3")
        print("<Response [429]>': Cannot handle too may requests, sleeping for 20 seconds")
        print("------------------------------------------------------------------------------")
        #time.sleep(20)
        countdown(20)
        os.system('cls')
        print("------------------------------------------------------------------------------")
        print(str(len(counter2))+" Companies extracted")
        print("------------------------------------------------------------------------------")
        quick_ratio_page = s.get(quick_ratio_url)
    
    quick_ratio_page_data = BeautifulSoup(quick_ratio_page.content, 'html.parser')#parse the webpage as html

    quick_ratio_items = quick_ratio_page_data.select(".name") #get the quick ratio names from webpage above (with html tags)
    quick_ratio_values = quick_ratio_page_data.select(".number") #get the quick ratio values from webpage above (with html tags)
    if(quick_ratio_items==[]):
        print("------------------------------------------------------------------------------------------------")
        print("No quick/custom ratios were configured in Screener for {}.Exiting the program".format(company))
        print("------------------------------------------------------------------------------------------------")
        sys.exit()

    profitloss_section=parse_web_page_data.select(".ranges-table") #get profitloss section from webpage

    cpg_ratio_sec = profitloss_section[1] #get CPG section
    cpg_ratio_dataset=cpg_ratio_sec.find_all("td") #extract CPG data with html tags

    cpg_ratio_items=[] #intialize empty array

    for item in cpg_ratio_dataset: # loop to extract data and remove html tags from each item
        item=(item.string).replace('%','')
        if(item==''):
            item='0'
        cpg_ratio_items.append(item)

    cpg_ratios = [cpg_ratio_items[0],cpg_ratio_items[2],cpg_ratio_items[4],cpg_ratio_items[6]] #get cpg ratios from items array
    cpg_values = [float(cpg_ratio_items[1]),float(cpg_ratio_items[3]),float(cpg_ratio_items[5]),float(cpg_ratio_items[7])] #get cpg values from items array

    default_ratio_items.extend(quick_ratio_items)
    default_ratio_values.extend(quick_ratio_values)
    #actual_ratios = ['Trigger']
    
    actual_ratios = [] # empty array for ratios labels
    for item in default_ratio_items:
        item=(item.string).replace('\n','').strip()
        actual_ratios.append(item)
    actual_ratios.extend(cpg_ratios)
    #actual_values=[company]
     
    actual_values=[] # empty array for values
    for value in default_ratio_values:
        value = value.string
        if ((value==None) or (value=='')):
            value = '0.00'
        value = value.replace('\n','').strip()
        value = value.replace(',','')
        actual_values.append(float(value))
    actual_values.extend(cpg_values)
    
    #print(actual_values)
              
    if "High / Low" in actual_ratios:
        HL_index = actual_ratios.index("High / Low")
        #print("Item no. {} contains H/L".format(HL_index+1) )
        #print(actual_values[HL_index],actual_values[HL_index+1])
        actual_values.insert(HL_index,str(actual_values[HL_index])+"/"+str(actual_values[HL_index+1]))
        actual_values.pop(HL_index+1)
        actual_values.pop(HL_index+1)
        
    ratio_value = {} # initializing empty dictionary
    for i in range(0,len(actual_ratios)):
        for ratio in required_ratios:
            if (ratio==actual_ratios[i]):
                ratio_value.update({required_ratios[ratio]:actual_values[i]}) # creating a ratio and value dictionary
    
    if(len(actual_values)==len(actual_ratios)):
        for x in required_ratios.values():
            if x not in ratio_value.keys():
                print("------------------------------------------------------------------------------------------------")
                print(x+"ratio not available")
                print("Check if you have configured the correct ratios in screener")
                print("------------------------------------------------------------------------------------------------")
                sys.exit()
            else:
                v1=float(ratio_value['10 Yrs'])
                v2=float(ratio_value['5 Yrs'])
                v3=float(ratio_value['3 Yrs'])
                v4=float(ratio_value['TTM'])
                v=float(ratio_value['PE'])

                cal_ratio={10.0 : 0 if v1==0 else v/v1
                           ,5.0 : 0 if v2==0 else v/v2
                           ,3.0 : 0 if v3==0 else v/v3
                           ,'TTMc': 0 if v4==0 else v/v4}
        ratio_value.update(cal_ratio)
        ratio_value_1up.append({company:ratio_value}) # list of dictionaries of ratios by companies)
    else:
        print("------------------------------------------------------------------------------------------------")
        print("Something wrong with values extracted.Kindly check screener")
        print("------------------------------------------------------------------------------------------------")
    print(len(counter1),' : ',str(len(counter2))+" Companies extracted",company)
os.system('cls')
print("------------------------------------------------------------------------------------------------")    
print(str(len(counter2))+" Companies extracted",company)
print('Data could not be found for these companies' + str(nodata))    
print("------------------------------------------------------------------------------------------------")
ratio_value_final=ratio_value_1up # list of dictionaries of ratios by companies)

for x in ratio_value_final:
    company=list(x.keys())[0]
    values=list(x.values())[0]
    actual_file.loc[company,list(values.keys())]=list(values.values()) # for each row label (Trigger) update the values
    actual_file.loc[company,'Last Updated']=dt.date(dt.today())# for each row label (Trigger) update today's date

actual_file = actual_file.reset_index()

actual_file=actual_file.drop(columns='index')

new_headers=['','','','','','','','','','','','','','CPG','CPG','CPG','CPG','PE/CPG','PE/CPG','PE/CPG','PE/CPG','','','','']

actual_file.loc[-1]=actual_file.columns
actual_file.index = actual_file.index + 1
actual_file = actual_file.sort_index()
actual_file.columns=new_headers

# validation if file is open
try:
    actual_file.to_excel(file_path,index=False)
    print("------------------------------------------------------------------------------------------------")
    print("Extraction completed and file saved here: " + file_path)
    print("------------------------------------------------------------------------------------------------")
except PermissionError:
    print("File with the same name is open, please close the file.You have 20 seconds")
    #time.sleep(10)
    countdown(60)
    try:
        actual_file.to_excel(file_path,index=False)
        print("------------------------------------------------------------------------------------------------")
        print("Extraction completed and file saved here: " + file_path)
        print("------------------------------------------------------------------------------------------------")
    except PermissionError:
        print("------------------------------------------------------------------------------------------------")
        print("File is still open. Exiting the program.........")
        print("------------------------------------------------------------------------------------------------")

end = time.time()
total_execution_time = (end - start)/60
print("------------------------------------------------------------------------------------------------")
print("Process Completed. Total time taken {} mins.".format(total_execution_time))
print("------------------------------------------------------------------------------------------------")
sys.exit

