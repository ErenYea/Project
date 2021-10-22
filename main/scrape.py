from selenium import webdriver
from selenium.webdriver.chrome import options
import time,csv
import re
import main.constants as const
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import win32com.client as win32

import pythoncom,os

class Scrape(webdriver.Chrome):
    options = options.Options()
    options.headless = False
    options.add_argument("window-size=1200x600")
    options.add_experimental_option('useAutomationExtension', False)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # initializing the webdriver instance

    def __init__(self, ):
        super(Scrape, self).__init__(options=self.options)
        self.result = {}
        self.results = {}
        self.state = None
        self.states = {}
        self.city = None
        self.cities = []
        self.date_from = None
        self.date_to = None
        self.count = 1
        self.absPath = os.path.abspath('results.xlsx')
       
        pythoncom.CoInitialize()
        self.dff = pd.read_excel('results.xlsx')
       
        self.ExcelApp = win32.gencache.EnsureDispatch("Excel.Application")
        self.ExcelApp.Visible = True
        self.ExcelApp.WindowState = win32.constants.xlMaximized
       
        self.headers = ['State', 'City', 'Range of Dates from:', 'Range of Dates to:', 'FULL NAME OF THE DECEASED PERSON WITHOUT COMMAS', 'FULL NAME OF THE DECEASED PERSON WITH COMMAS', 'YEAR OF BIRTH', 'YEAR OF DEATH', 'DATE OF DEATH', 'Funeral Home Name',
                            'Funeral Home Street Address', 'Funeral Home City', 'Funeral Home State', 'Funeral Home ZIP Code', 'Upcoming Service Name', 'Upcoming Service Date', 'Upcoming Service City', 'List of Next of Kin', "Link to the deceased person's obituary"]

        if len(self.dff) == 0:


            self.wb = self.ExcelApp.Workbooks.Open(self.absPath)
            
            self.ws = self.wb.Worksheets("Sheet1")
           
            
            for i in range(1,len(self.headers)+1):
                self.ws.Cells(1,i).Value = self.headers[i-1]
                self.ws.Cells(1,i).Font.Name = 'Verdana'
                self.ws.Cells(1,i).Font.Size = 13
                self.ws.Cells(1,i).Font.Bold = True
                
        else:
            self.wb = self.ExcelApp.Workbooks.Open(self.absPath)
            
            self.ws = self.wb.Worksheets("Sheet1")    
            for i in range(1,len(self.headers)+1):
                self.count = 2
                self.ws.Cells(1,i).Value = self.headers[i-1]
                self.ws.Cells(1,i).Font.Name = 'Verdana'
                self.ws.Cells(1,i).Font.Size = 13
                self.ws.Cells(1,i).Font.Bold = True
                for j in self.dff[self.headers[i-1]]:
                    if str(j) == 'nan':
                        self.ws.Cells(self.count,i).Value = "-"
                    else:
                        self.ws.Cells(self.count,i).Value = str(j)
                    self.count += 1
                

        self.ws.Columns.AutoFit()
        self.ws.Rows.AutoFit()
        
        self.implicitly_wait(const.IMPLICIT_WAIT)
        
        self.keywords = pd.read_csv('keywords.csv')
        with open('file.csv','r') as f:
            csv_reader = csv.reader(f)
            self.csv_list = []
            
            for i in csv_reader:
                print(i)
                if len(i) == 0:
                    continue
                self.csv_list.append(i[0])
        print(self.csv_list) 

    # Loading the frist page
    def land_on_first_page(self):
        self.get(const.BASE_URL)

    def click_on_popup(self):

        btn = self.find_element_by_xpath(
            "//div[@class='fc-dialog-container']/div/div[2]/div[2]/button")

        print(btn)

        btn.click()
    
    def ad_pop_up(self):
        element = WebDriverWait(self, 30).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="fEy1Z2XT "]/div/div/div/div[3]/span/button')))
        # element = self.find_element_by_xpath('//div[@class="fEy1Z2XT "]/div/div/div/div[3]/span/button')
        element.click()
         

    # Selecting the country
    def select_contry(self):
        select = Select(self.find_element_by_id(
            'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_ddlCountry'))
        select.select_by_visible_text('United States')

    # Getting the names of state

    def get_states(self):
        states = self.find_elements_by_xpath(
            "//select[@id='ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_ddlState']/option")
        for i in states:
            print(f"Value: {i.get_attribute('value')} , Text: {i.text}")
            self.states[i.get_attribute('value')] = i.text

    # Takingt the input of state
    def input_state(self, state=''):
        try:
            select = Select(self.find_element_by_id(
                'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_ddlState'))
        except:
            select = Select(self.find_element_by_xpath('//select[@name="ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$uxSearchWideControl$ddlState"]'))

        if state == '':
            
            select.select_by_value('57')
            self.state = self.states['57']
            
        else:
            
            select.select_by_visible_text(state)
            self.state = state

     # selecting the keywords
    def keyword(self, keyword=''):
        self.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        
        button = WebDriverWait(self, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_txtKeyword')))
        print(button)
        
        key = self.find_element_by_xpath(
            '//div[@class="trKeyword"]/input')
        
        self.city = keyword
        
        try:

            key.clear()
        except:
            button = WebDriverWait(self, 10).until(EC.element_located_to_be_selected((By.ID, 'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_txtKeyword')))
            print(button)
            ActionChains(self).move_to_element(key).click(key).perform()
            key.clear()
        key.send_keys(keyword)

    # Selecting the date

    def select_date(self):
        select = Select(self.find_element_by_id(
            'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchWideControl_ddlSearchRange'))
        select.select_by_value('88888')

    # Selcting the date range
    def date_range(self, date_from='02/12/2020', date_to='02/12/2021'):
        div_tag_for_date = self.find_elements_by_class_name('DateValue')
        self.date_from = date_from
        self.date_to = date_to
        print(len(div_tag_for_date))
        date_from_tag = div_tag_for_date[0].find_element_by_tag_name('input')
        date_to_tag = div_tag_for_date[1].find_element_by_tag_name('input')
        date_from_tag.clear()
        date_to_tag.clear()
        date_from_tag.send_keys(date_from)
        date_to_tag.send_keys(date_to)

    # Clicking on search button
    def search(self):
        search = self.find_element_by_link_text("Search")
        search.click()

    # testing the condtition of result
    def get_result(self):
        try:

            txt = self.find_element_by_xpath("//div[@class='InlineTotalCountText']").text
            lst = [int(x) for x in txt.split() if x.isdigit()]
            print(max(lst))
            if max(lst) <= 10:
                return "less than 10"
        except:
            try:
                result = self.find_element_by_class_name('RefineMessage').text
                print(result)
                if '1000+' in result:
                    return True
                elif 'did not find any obituaries' in result:
                    return 'Didnot'
                else:
                    return False
            except:
                return False

    def click_all_results(self):
        try:
            result = self.find_element_by_class_name('RefineMessage').text
            if 'View all results.' in result:
                self.find_element_by_id(
                    'ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_uxSearchLinks_ViewAllLink').click()
        except:
            pass

    # scrolling down the window to show all the results
    def scrolldown(self):
        # Get scroll height
        last_height = self.execute_script("return document.body.scrollHeight")
        print(f"last_height {last_height}")

        while True:
            # Scroll down to bottom
            self.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(const.SCROLL_PAUSE_TIME)

            # Calculate new scroll height and compare with last scroll height
            new_height = self.execute_script(
                "return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

    def result_to_csv(self, name='result.csv'):
        results = self.find_elements_by_xpath('//div[@class="mainScrollPage"]')
        self.result = {}
        for i in results:
            a = i.find_elements_by_class_name('entryContainer')
            print(len(a))
            for j in a:
                s = j.find_element_by_class_name("obitName")
                h = s.find_element_by_tag_name('a')
                
                if h.get_attribute('href') in self.csv_list:
                    continue
                print(f"TExt: {s.text}  link: {h.get_attribute('href')}")
                
                self.result[s.text] = h.get_attribute('href')
                print("\n")
        

    def read_result(self, key):
        url = self.result[key]
        driver = webdriver.Chrome(options=self.options)
        
        print(
            f"----------------- Extracting Data about {key} -----------------")
        print('')
        try:
            driver.get(url)
            driver.implicitly_wait(30)
            try:
                self.ad_pop_up()
            except Exception as e:
                print(e)
            if driver.current_url in list(self.csv_list):
            
                return
            if "legacy" in driver.current_url :
                try:
                    self.ad_pop_up()
                except Exception as e:
                    print(e)
                try:
                    para = driver.find_element_by_xpath(
                        "//div[@data-component='ObituaryParagraph']").text.split('.')

                    try:
                        dob = driver.find_element_by_xpath(
                            "//div[@class='Box-sc-5gsflb-0 iobueB']/div/div/div/div").text
                        dod = driver.find_element_by_xpath(
                            "//div[@class='Box-sc-5gsflb-0 iobueB']/div/div[2]/div/div").text

                    except:
                        dob = '-'
                        dod = '-'

                    try:
                        funeral_home_list = driver.find_element_by_xpath(
                            "//div[@class='Box-sc-5gsflb-0 iobueB']/div[2]/div").text.split('\n')
                        funeral_home_name = funeral_home_list[1]
                        funeral_home_street = funeral_home_list[2]
                        funeral_home_city = funeral_home_list[3].split(',')[0]
                        funeral_home_state = funeral_home_list[3].split(',')[1]
                        funeral_home_zipcode = '-'

                    except:
                        try:
                            funeral_home_list = driver.find_element_by_xpath(
                                "//div[@class='Box-sc-5gsflb-0 iobueB']/div[2]/div").text.split('\n')
                            funeral_home_name = funeral_home_list[1]
                            funeral_home_street = funeral_home_list[2]
                            funeral_home_city = funeral_home_list[3].split(',')[0]
                            funeral_home_state = funeral_home_list[3].split(',')[1]
                            funeral_home_zipcode = '-'
                        except:
                            funeral_home_name = '-'
                            funeral_home_street = '-'
                            funeral_home_city = '-'
                            funeral_home_state = '-'
                            funeral_home_zipcode = '-'
                    try:
                        date_of_death = re.findall(
                            "\w+.\s+\d{1,2},\s+\d{4}", para[0])[0]
                    except:
                        date_of_death = '-'

                    TITLE = r"(?:[A-Z][a-z]*\.\s*)?"
                    NAME1 = r"[A-Z][a-z]+,?\s+"
                    MIDDLE_I = r"(?:[A-Z][a-z]*\.?\s*)?"
                    NAME2 = r"[A-Z][a-z]+"
                    res = re.findall(TITLE + NAME1 + MIDDLE_I + NAME2, para[0])
                    if 'In Loving Memory' in res[0]:
                        full_name = res[1]
                    else:
                        full_name = res[0]
                    if ',' in full_name:
                        full_name_with_commas = full_name
                        full_name_without_commas = ''
                    else:
                        full_name_with_commas = ''
                        full_name_without_commas = full_name

                    try:
                        upcoming_service_list = driver.find_elements_by_xpath(
                            "//div[@class='Box-sc-5gsflb-0 bQzMjo']/div[2]/div[@class='Box-sc-5gsflb-0 kwgeEM']")[0].text.split('\n')
                        if 'Plant Memorial Trees' in upcoming_service_list[-1]:
                            upcoming_service_month = ''
                            upcoming_service_day = '-'
                            upcoming_service_name = upcoming_service_list[-1]
                        else:
                            upcoming_service_month = upcoming_service_list[0]
                            upcoming_service_day = upcoming_service_list[1]
                            upcoming_service_name = upcoming_service_list[2]
                    except:
                        upcoming_divs = driver.find_elements_by_xpath(
                            "//div[@class='Box-sc-5gsflb-0 bQzMjo']/div[2]/div[@class='Box-sc-5gsflb-0 irxurr']")
                        upcoming_service_month = []
                        upcoming_service_day = []
                        upcoming_service_name = []

                        for i in upcoming_divs:
                            j = i.text.split('\n')
                            if 'Plant Memorial Trees' in j[-1]:
                                upcoming_service_month.append('-')
                                upcoming_service_day.append('-')
                                upcoming_service_name.append(j[-1])
                            else:
                                upcoming_service_month.append(j[0])
                                upcoming_service_day.append(j[1])
                                upcoming_service_name.append(j[2])

                    upcoming_service_date = ''
                    try:

                        for i in range(0, len(upcoming_service_month)):
                            if i == len(upcoming_service_month)-1:
                                upcoming_service_date += f'{upcoming_service_month[i]}-{upcoming_service_day[i]}'
                            else:
                                upcoming_service_date += f'{upcoming_service_month[i]}-{upcoming_service_day[i]}, '

                        if len(upcoming_service_name) == 1:
                            upcoming_service_name = upcoming_service_name[0]
                        else:
                            upcoming_service_names = upcoming_service_name
                            upcoming_service_name = ''
                            for i in range(0, len(upcoming_service_names)):
                                if i == len(upcoming_service_names)-1:
                                    upcoming_service_name += f'{upcoming_service_names[i]}'
                                else:
                                    upcoming_service_name += f'{upcoming_service_names[i]}, '

                    except:
                        upcoming_service_name = ''
                        upcoming_service_date = ''

                    
                    lst = []
                    for i in range(len(self.keywords)):
                        for j in para:
                            if self.keywords.loc[i,'Keywords'] in j:
                                if j in lst:
                                    continue
                                lst.append(j)

                    lonok = ''
                    for i in lst:
                        lonok += i

                    rows = {'State': self.state, 'City': self.city, 'Range of Dates from:': self.date_from, 'Range of Dates to:': self.date_to, 'FULL NAME OF THE DECEASED PERSON WITHOUT COMMAS': full_name_without_commas, 'FULL NAME OF THE DECEASED PERSON WITH COMMAS': full_name_with_commas, 'YEAR OF BIRTH': dob, 'YEAR OF DEATH': dod, 'DATE OF DEATH': date_of_death, 'Funeral Home Name': funeral_home_name,
                            'Funeral Home Street Address': funeral_home_street, 'Funeral Home City': funeral_home_city, 'Funeral Home State': funeral_home_state, 'Funeral Home ZIP Code': funeral_home_zipcode, 'Upcoming Service Name': upcoming_service_name, 'Upcoming Service Date': upcoming_service_date, 'Upcoming Service City': funeral_home_city, 'List of Next of Kin': lonok, 'Link to the deceased person': url}


                    print(rows)

                    self.count += 1
                    for index,key in enumerate(rows):
                        self.ws.Cells(self.count,index+1).Value = rows[key]
                    
                        
                    self.ws.Columns.AutoFit()
                    self.ws.Rows.AutoFit()
                    with open('file.csv','a') as f:
                        csv_writer = csv.writer(f)
                        csv_writer.writerow([url])
                   
                except Exception as e:
                    print(e)
                    driver.close()
                    with open('file.csv','a',newline="") as f:
                        csv_writer = csv.writer(f)
                        csv_writer.writerow([url])
                    
                    return
        except Exception as e:
            print(e)
            print('Url denied')
        driver.close()

    def runscrapper(self):

        for key in self.result:
            self.read_result(key)
        self.wb.Close(True) # save the workbook
        
        
        