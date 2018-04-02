import time,datetime,calendar
import ezodf,pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from time import strptime
from pandas import ExcelWriter
from pandas import ExcelFile

path='C:\\Users\\Samridi\\Downloads\\sample_records.ods'
def getpd(path):
    doc = ezodf.opendoc(path)
    sheet = doc.sheets[0]
    df_dict = {}
    for i, row in enumerate(sheet.rows()):
        # row is a list of cells
        # assume the header is on the first row
        if i == 0:
            # columns as lists in a dictionary
            df_dict = {cell.value:[] for cell in row}
            # create index for the column headers
            col_index = {j:cell.value for j, cell in enumerate(row)}
            continue
        for j, cell in enumerate(row):
            # use header instead of column index
            df_dict[col_index[j]].append(cell.value)
    df = pd.DataFrame(df_dict)
    return(df)
toaddarr=[]
ndf=getpd(path)
df=ndf.dropna()
df=ndf[['trackingnumber','shipdate','country','shipcountry','pincode','shippincode','servicetype']].copy()
count=0
df=df.head(50)
for index, row in df.iterrows():
    try:
        inDate=row['shipdate']  #get input
        #getting date from string to datetime format
        d = datetime.datetime.strptime(inDate, '%Y-%m-%d %H:%M:%S')
        shipdate='/'.join(str(x) for x in (d.day, d.month, d.year))
        print("shipdate - "+str(inDate))
        oo=[int(d.year),int(d.month),int(d.day)]

        #getting next day no
        dno=(datetime.datetime(oo[0],oo[1],oo[2])).weekday()
        def nextdate(no):
            d = datetime.date.today()
            while d.weekday() != no:
                 d += datetime.timedelta(1)
            return d
        dno=nextdate(dno)
        dno=str(dno.strftime("%m-%d-%Y")).replace('-','/')
        print("next date with same dno - "+str(dno))

        # scraping for data
        chromedriver= 'D:\\webdrivers\\chromedriver'
        options = Options()
        options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        browser = webdriver.Chrome(chromedriver,chrome_options=options)
        browser.get('https://www.fedex.com/ratefinder/home')

        fromcountry = browser.find_element_by_name("origCountry")
        destCountry = browser.find_element_by_name("destCountry")


        el = browser.find_element_by_id('origCountryId')
        for option in el.find_elements_by_tag_name('option'):
            fromctry=row['country']#'IN'  #get input
            if option.get_attribute("value") == fromctry:
                option.click()
                break

        ll = browser.find_element_by_id('destCountryId')
        for option in ll.find_elements_by_tag_name('option'):
            toctry=row['shipcountry']#'IN' #get input
            if option.get_attribute("value") == toctry:
                option.click()
                break
        try: #try zipcode else by city
            zplid = browser.find_element_by_id("origZipId")
            destzpid= browser.find_element_by_id("destZipId")
            zpin=str(row['pincode'])
            zpin= zpin.replace('.0','')
            destpin=str(row['shippincode'])
            destpin= destpin.replace('.0','')
            #send keys
            zplid.send_keys(zpin)  #get input for pin
            destzpid.send_keys(destpin)#get input for pin
        except:
            try:
                ogcity= browser.find_element_by_id("origCityId")
                for option in ogcity.find_elements_by_tag_name('option'):
                    toctry=row['city'] #get shipcity input
                    if option.get_attribute("value") == toctry:
                        option.click()
                        break
                descity= browser.find_element_by_id("destCityId")
                for option in descity.find_elements_by_tag_name('option'):
                    toctry=row['shipcity'] #get city input
                    if option.get_attribute("value") == toctry:
                        option.click()
                        break
            except:
                pass
        no = browser.find_element_by_id("NumOfPackages")
        weigh=browser.find_element_by_id("totalPackageWeight")
        no.clear()
        no.send_keys("1")
        weigh.send_keys("10")
        
        pk = browser.find_element_by_name('receivedAtCode')
        for option in pk.find_elements_by_tag_name('option'):
            if option.text == 'Drop off at FedEx location':
                option.click()
                break

        browser.execute_script("document.getElementById('shipCalendarDate._date').value ='"+dno+"' ;")
        browser.find_element_by_id("ttTime").click()
        try:
            ll = browser.find_element_by_name('shipmentPurpose')
            for option in ll.find_elements_by_tag_name('option'):
                if option.text == 'Personal (Not sold)':
                    option.click()
                    break

            

            ll = browser.find_element_by_name('freightOnValue')
            for option in ll.find_elements_by_tag_name('option'):
                if option.text == 'Own risk':
                    option.click()
                    break
        except:
            pass
        try:
            tiv = browser.find_element_by_name("customsValue")
            tiv.send_keys("10")
        except:
            print("failed to set custom value")
        
        lfg = browser.find_element_by_name('packageForm.packageList[0].packageType')
        for option in lfg.find_elements_by_tag_name('option'):
            if option.text == "Your Packaging":
                option.click()
                break
        browser.find_element_by_class_name("buttonpurple").click()
        
        svtype=str(row['servicetype'])
        print(svtype)
        servicetype=svtype  #get service type
        arr=[]
        flag=0
        if browser.find_element_by_xpath("//td[@style='border-right: 1px solid #CFCFCF; vertical-align: middle;']//font[contains(@id,'"+servicetype+"')]"):
            ao=browser.find_element_by_xpath("//td[@style='border-right: 1px solid #CFCFCF; vertical-align: middle;']//font[contains(@id,'"+servicetype+"')]").text
            if "End of day" in ao:
                try:
                    flag=int(ao[11])
                except:
                    flag=2
            else:
                arr=str(ao).replace(',','').split(' ')
                arr=arr[1:4]
        if flag==0:
            arr[0]=str(strptime(arr[0],'%b').tm_mon)
            finaldate='/'.join(arr)
            print("date calculated from dno - "+str(finaldate))
            date_format = "%m/%d/%Y"
            a = datetime.datetime.strptime(dno, date_format)
            b = datetime.datetime.strptime(finaldate, date_format)
            delta = b - a
            print("No of days taken = "+str(delta.days))
            #calcuation for the estimated date
            erp=d+ datetime.timedelta(days=int(delta.days))
            print("estimation date= "+str(erp))
            toaddarr.append(str(erp))
            print("------------------------------------------")
        else:
            erp=d+ datetime.timedelta(days=int(flag))
            print("estimation date= "+str(erp))
            toaddarr.append(str(erp))
            print("------------------------------------------")
        browser.quit()
    except:
        print("error case failed")
        browser.quit()
        count+=1
        toaddarr.append("error occured")
        print("------------------------------------------")
        pass
print("no of exceptions= "+str(count))
print(df.shape)
df['estimated date']=toaddarr
print(df.shape)
writer = ExcelWriter('fedex_estimate_date_cal.xlsx')
df.to_excel(writer,'Sheet1',index=False)
writer.save()
