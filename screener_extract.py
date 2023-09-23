import requests
import os
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime as dt
import sys
import time
import openpyxl

start = time.time()

def countdown(time_sec):
    while time_sec:
        mins, secs = divmod(time_sec, 60)
        timeformat = '{:02d}:{:02d}'.format(mins, secs)
        print(timeformat, end='\r')
        time.sleep(1)
        time_sec -= 1
    print("done")

def screener_login(url,loginextn,userid,user_password):
    s=requests.Session()
    main_url = url
    login_sup = loginextn
    home_url = os.path.join(main_url,login_sup)
    resp=s.get(home_url)
    if 'csrftoken' in resp.cookies:
        # Django 1.6 and up
        csrftoken = resp.cookies['csrftoken']
    else:
        # older versions
        csrftoken = resp.cookies['csrf']
    login_data = dict(username=userid, password=user_password, csrfmiddlewaretoken=csrftoken, next='/')
    r = s.post(home_url, data=login_data, headers=dict(Referer=home_url))
    print(s)
    return s

required_ratios = {'Market Cap':'M Cap'
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

columns_list = ['Trigger','M Cap','Sh Pr','PE','P/B','ROCE', 'ROE','ROA', 'DTE','OPM', 'EY','PEG', 'DY','10 Yrs','5 Yrs','3 Yrs','TTM',10.0,5.0,3.0,'TTMc']

main_url = 'https://www.screener.in/'
login_sup = 'login/?'

print("Enter your file path:")
file_path = input()
print("->"+ file_path)


try:
    actual_file = pd.read_excel(file_path)
    print("File successfuly uploaded.")
except FileNotFoundError:
    print("File unavailable.Please place the file.")
    sys.exit()
except ValueError:
    print("Unrecognized file format.Please provide the correct file.")
    sys.exit()

if(list(actual_file.loc[0][0:21])==columns_list):
   print("No issues with file columns.Proceeding with data extraction")

else:
   print("File column names are incorrect, please upload the correct file and proceed")
   sys.exit()


print("Enter your Screener User ID: ")
user_id = input()
print("->"+ user_id)

print("Enter your Screener Password: ")
pass_word = input()
print("->"+ pass_word)

actual_file.columns=actual_file.loc[0]

actual_file = actual_file.loc[:, actual_file.columns.notnull()]

actual_file.index=[actual_file['Trigger']]

actual_file.index.names = ['index']

actual_file.drop(actual_file.index[:1],inplace=True)

search_list = list(actual_file["Trigger"])
if(search_list==[]):
    print('Search Company Empty in file. Please upload the file with data.')
    sys.exit()
else:
    search_list=[x for x in search_list if x==x]
    print("{} Companies to extract".format(len(search_list)))

s=screener_login(main_url,login_sup,user_id,pass_word)

ratio_value_list=[]
counter1=[]
counter2=[]
nodata=[]
ratio_value_1up=[]

for company in search_list:
    if (len(counter1)==9):
        counter1=[]
        end = time.time()
        execution_time = (end - start)/60
        print("Time taken: {} mins".format(execution_time))
        print("Started..........")
    url_extn = 'company/'+company+'/consolidated/'
    get_webpage=s.get(os.path.join(main_url,url_extn))
    while(str(get_webpage)=='<Response [429]>'):
        print("print-1")
        print("<Response [429]>': Cannot handle too may requests, wait for 20 seconds")
        #time.sleep(20)
        countdown(20)
        print("Started..........")
        get_webpage=s.get(os.path.join(main_url,url_extn))
    if(str(get_webpage)=='<Response [404]>'):
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
    
    #quick ratios
    datawarehouseid_tag=parse_web_page_data.main.div
    datawarehouseid = datawarehouseid_tag['data-warehouse-id'] #get warehouse id
    
    if (datawarehouseid=='None'):
        url_extn = 'company/'+company+'/'
        get_webpage=s.get(os.path.join(main_url,url_extn))
        if(str(get_webpage)=='<Response [404]>'):
            print(company+" warehouseid could not found to extract quick ratios")
            nodata.append(company)
            continue
        while(str(get_webpage) =='<Response [429]>'):
            print("print-2")
            print("<Response [429]>': Cannot handle too may requests, wait for 20 seconds")
            #time.sleep(20)
            countdown(20)
            print("Started..........")    
            get_webpage=s.get(os.path.join(main_url,url_extn))
    
    parse_web_page_data = BeautifulSoup(get_webpage.content, 'html.parser')
    datawarehouseid_tag=parse_web_page_data.main.div
    datawarehouseid = datawarehouseid_tag['data-warehouse-id'] #get warehouse id
    
    quick_ratio_url = os.path.join(main_url,'api/company/'+datawarehouseid+'/quick_ratios/')#create api url
    quick_ratio_page = s.get(quick_ratio_url)#get the api webpage
    
    while(str(quick_ratio_page) =='<Response [429]>'):
        print("print-3")
        print("<Response [429]>': Cannot handle too may requests, sleeping for 20 seconds")
        #time.sleep(20)
        countdown(20)
        print("Started..........")    
        quick_ratio_page = s.get(quick_ratio_url)
    
    quick_ratio_page_data = BeautifulSoup(quick_ratio_page.content, 'html.parser')#parse the webpage as html

    quick_ratio_items = quick_ratio_page_data.select(".name") #get the quick ratio names from webpage above (with html tags)
    quick_ratio_values = quick_ratio_page_data.select(".number") #get the quick ratio values from webpage above (with html tags)
    if(quick_ratio_items==[]):
        print("No quick/custom ratios were configured in Screener for {}.Exiting the program".format(company))
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
    cpg_values = [cpg_ratio_items[1],cpg_ratio_items[3],cpg_ratio_items[5],cpg_ratio_items[7]] #get cpg values from items array
    default_ratio_items.extend(quick_ratio_items)
    default_ratio_values.extend(quick_ratio_values)
    #actual_ratios = ['Trigger']
    actual_ratios = []
    for item in default_ratio_items:
        item=(item.string).replace('\n','').strip()
        actual_ratios.append(item)
    actual_ratios.extend(cpg_ratios)
    #actual_values=[company]
    actual_values=[]
    for value in default_ratio_values:
        value = value.string
        if ((value==None) or (value=='')):
            value = '0.00'
        value = value.replace('\n','').strip()
        actual_values.append(value)
    actual_values.extend(cpg_values)
    if "High / Low" in actual_ratios:
        HL_index = actual_ratios.index("High / Low")
        #print("Item no. {} contains H/L".format(HL_index+1) )
        #print(actual_values[HL_index],actual_values[HL_index+1])
        actual_values.insert(HL_index,actual_values[HL_index]+"/"+actual_values[HL_index+1])
        actual_values.pop(HL_index+1)
        actual_values.pop(HL_index+1)
    ratio_value = {}
    for i in range(0,len(actual_ratios)):
        for ratio in required_ratios:
            if (ratio==actual_ratios[i]):
                ratio_value.update({required_ratios[ratio]:actual_values[i]})
    
    if(len(actual_values)==len(actual_ratios)):
        for x in required_ratios.values():
            if x not in ratio_value.keys():
                print(x+"ratio not available")
                print("Check if you have configured the correct ratios in screener")
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
        ratio_value_1up.append({company:ratio_value})
    else:
        print("Something wrong with values extracted.Kindly check screener")
    print(len(counter1),str(len(counter2))+" Companies extracted",company)
    
print('Data could not be found for these companies' + str(nodata))    
ratio_value_final=ratio_value_1up


for x in ratio_value_final:
    company=list(x.keys())[0]
    values=list(x.values())[0]
    actual_file.loc[company,list(values.keys())]=list(values.values())
    actual_file.loc[company,'Last Updated']=dt.date(dt.today())

actual_file = actual_file.reset_index()

actual_file=actual_file.drop(columns='index')

new_headers=['','','','','','','','','','','','','','CPG','CPG','CPG','CPG','PE/CPG','PE/CPG','PE/CPG','PE/CPG','','','','']

actual_file.loc[-1]=actual_file.columns
actual_file.index = actual_file.index + 1
actual_file = actual_file.sort_index()
actual_file.columns=new_headers

try:
    actual_file.to_excel(file_path,index=False)
    print("Extraction completed and file saved here: " + file_path)
except PermissionError:
    print("File with the same name is open, please close the file.You have 20 seconds")
    #time.sleep(10)
    countdown(20)
    try:
        actual_file.to_excel(file_path,index=False)
        print("Extraction completed and file saved here: " + file_path)
    except PermissionError:
        print("File is still open. Exiting the program.........")
    
    
end = time.time()
total_execution_time = (end - start)/60
print("Process Completed. Total time taken {} mins.".format(total_execution_time))
sys.exit
