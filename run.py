from types import ClassMethodDescriptorType

from pandas.core.indexing import is_nested_tuple
from main.scrape import Scrape
import pandas as pd
import pythoncom

pythoncom.CoInitialize()

inst = Scrape()
inst.land_on_first_page()

try:
    inst.ad_pop_up()
except Exception as e:
    print(e)
    
inst.select_contry()

choice = input("Type 'Y' to chose by state or Type 'N' for chose by city: ").lower()
if choice == 'n':
    inst.select_date()
    print('\n')
    date_from = input("Enter the date from in format mm/dd/yyyy: ")
    print('\n')
    date_to = input("Enter the date to in format mm/dd/yyyy: ")
    inst.date_range(date_from=date_from,date_to=date_to)
    df = pd.read_csv('city.csv')
    for i in range(len(df)) :
        print(df.loc[i, "City"],' , ' ,df.loc[i, "State"])
        inst.input_state(state=df.loc[i, "State"]) 
        keyword = df.loc[i, "City"]
        inst.keyword(keyword=keyword)
        inst.search()
        
            
        try:
            inst.ad_pop_up()
        except Exception as e:
            print(e)
        
        result_cond = inst.get_result()
        if result_cond == True:
            inst.count += 1
            rows = {'State': inst.state, 'City': inst.city, 'Range of Dates from:': inst.date_from, 'Range of Dates to:': inst.date_to, 'FULL NAME OF THE DECEASED PERSON WITHOUT COMMAS': "city gave a '1000+ Results'", 'FULL NAME OF THE DECEASED PERSON WITH COMMAS': '-', 'YEAR OF BIRTH': dob, 'YEAR OF DEATH': '-', 'DATE OF DEATH': '-', 'Funeral Home Name': '-',
                    'Funeral Home Street Address': '-', 'Funeral Home City': '-', 'Funeral Home State': '-', 'Funeral Home ZIP Code': '-', 'Upcoming Service Name': '-', 'Upcoming Service Date': '-', 'Upcoming Service City': '-', 'List of Next of Kin': '-', 'Link to the deceased person': '-'}
            for index,key in enumerate(rows):
                inst.ws.Cells(inst.count,index+1).Value = rows[key]
            inst.ws.Columns.AutoFit()
            inst.ws.Rows.AutoFit()
            continue
        elif result_cond == 'less than 10':
            inst.result_to_csv()
            inst.runscrapper()
        elif result_cond == 'Didnot':
            inst.count += 1
            rows = {'State': inst.state, 'City': inst.city, 'Range of Dates from:': inst.date_from, 'Range of Dates to:': inst.date_to, 'FULL NAME OF THE DECEASED PERSON WITHOUT COMMAS': "No Result", 'FULL NAME OF THE DECEASED PERSON WITH COMMAS': '-', 'YEAR OF BIRTH': dob, 'YEAR OF DEATH': '-', 'DATE OF DEATH': '-', 'Funeral Home Name': '-',
                    'Funeral Home Street Address': '-', 'Funeral Home City': '-', 'Funeral Home State': '-', 'Funeral Home ZIP Code': '-', 'Upcoming Service Name': '-', 'Upcoming Service Date': '-', 'Upcoming Service City': '-', 'List of Next of Kin': '-', 'Link to the deceased person': '-'}
            for index,key in enumerate(rows):
                inst.ws.Cells(inst.count,index+1).Value = rows[key]
            inst.ws.Columns.AutoFit()
            inst.ws.Rows.AutoFit()
            continue
        else:
            inst.click_all_results()
            inst.scrolldown()
            inst.result_to_csv()
            inst.runscrapper()

else:
    print('*******************************************************************')
    print('********************** Printing States ****************************')
    print('\n')
    inst.get_states()
    print('\n')
    print('********************************************************************')
    print('\n')

    state = input("Enter the name of state from above list or Just press Enter to Directly Use TEXAS : ")
    inst.input_state(state=state)
    print('\n')
    inst.select_date()
    while True:
        print('\n')
        date_from = input("Enter the date from in format mm/dd/yyyy: ")
        print('\n')
        date_to = input("Enter the date to in format mm/dd/yyyy: ")
        inst.date_range(date_from=date_from,date_to=date_to)
        inst.search()
        
        try:
            inst.ad_pop_up()
        except Exception as e:
            print(e)
       
        result_cond = inst.get_result()
        print(result_cond)
        if result_cond == False:
            break
        elif result_cond == 'Didnot':
            print('\n************************************************')
            print('***************** No result ********************')
            print('\n')
            print('Quitiing')
            quit()
        print("Result Exceeded from limit \nplease reduced the date range")
    inst.click_all_results()
    inst.scrolldown()
    inst.result_to_csv()
    inst.runscrapper()
print("Result got")
inst.close()
# inst.wb.SaveAs('cell_color.xlsx')
# inst.df.to_excel(r'result.xlsx', index = False)