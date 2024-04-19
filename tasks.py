from robocorp.tasks import task
from  RPA.Browser.Selenium import Selenium
from  Resources.Variables import *
from  RPA.Robocorp.utils import BuiltIn
from  RPA.Robocorp.WorkItems import WorkItems
from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
import requests
import json
import re
import shutil
import os
work_items_tools = WorkItems()
web_tools = Selenium()
built_in_tools = BuiltIn()
error_message = "Error - "
succes_message = "Succes -"
table_tools = Tables()
today_date = datetime.now()
excel_tools = Files()
locators = []
money_patterns = [r'$\d{0,3}.?\d+',r'$\d{0,3}\,\d{0,3}.\d{0,3}',r'\d+\bUSD\b',r'\d+\dollars\b']
media_folder = ".\Images"
try:
    os.mkdir(media_folder)
except:
    pass
with open("Locators\\all_locators.json", 'r') as file:
    # Load the JSON data
    locators = json.load(file)
def go_to_page():
    """
    Goes to home page
    """
    web_tools.open_available_browser(search_page)
    
def get_results():
    """
    Process all work items
    """
    search_params = work_items_tools.get_work_item_payload()
    built_in_tools.log_to_console(f"Payload to process {search_params}")
    found = search_item(search_topic=search_params["search_phrase"],result_category=search_params["news_categroy"],most_resent_data=int(search_params["data_time"]))
    
    if (found):
        fetch_more_result = scann_results(int(search_params["data_time"]))
        while (fetch_more_result):
            
            for _ in range( 3 ):
                try:
                    fetch_more_result = scann_results(int(search_params["data_time"]))
                    break
                except:
                    pass
            web_tools.click_element(locators["view_older_results_button"])
            
            

def search_item(search_topic:str, result_category:str,most_resent_data:int):
    """
    Searchs and detects if a results exist upon search
    """
    #check input  validity
    if len(search_topic) ==0 or len(result_category) == 0 or most_resent_data <0 :
        built_in_tools.log_to_console(f"Data not valid, Topic :{search_topic},category : {result_category}, months ago {most_resent_data}")
    web_tools.click_element(locator=locators["search_icon"])
    web_tools.input_text(locator=locators["search_box"],text=search_topic,clear=True) 
    web_tools.click_element(locator=locators["search_button"])
    if "No search results match the term " in web_tools.get_text(locators["No_availableresult"]):
        return False
    return True


def scann_results(months_ago): 
    """
    Reads all resulting rows and extract metadata
    months_ago : how old (in months) the results can be
    """
    #scan up to max number of rows

    #read data
    no_of_available_rows = web_tools.get_element_count(locator=locators["available_items_result"])
    results = []
    continue_to_fetch_more_results = True
    for index in range (1,no_of_available_rows+1):
        result = {"Title":"","Date":"","Description":"","Picture":"","PhrasesCount":"","MoneyAmountPresent":""}
        result["Title"] = web_tools.get_text(locator=locators["title_result_target"].replace("index",str(index)))
        try:
            image_download_url =  web_tools.get_element_attribute(locator=locators["image_result_targe"].replace("index",str(index)),attribute="src")
            result["Picture"] = image_download_url.split("/")[-1].split("jpg")[0]+"jpg"
            download_image(f"{media_folder}/{result['Picture']}",image_download_url)
        except:
            result["Picture"] = "Not available"
        result["Description"] = "Not available"
        result["PhrasesCount"] = len(result["Description"].split())
        # metadata  = web_tools.get_text(locator=locators["metada_result_target"].replace("index",str(index)))
         
        
        
        result["Date"]  = web_tools.get_text(locator=locators["date_result_target"].replace("index",str(index)))

   
        try:
            result["Date"]  = datetime.strptime(result["Date"] ,"%B %d, Y")
        except:
            result["Date"] = today_date
            

        result["MoneyAmountPresent"] = False
        for pattern in money_patterns:
            result["MoneyAmountPresent"]  = len(re.findall(pattern,result["Description"])) > 0

        if  not relativedelta(today_date ,result["Date"]).months<= months_ago:
            continue_to_fetch_more_results= False
        

        
            
        results.append(result)
    save_data_to_excel(results)
    return continue_to_fetch_more_results


def filter_results(result_category:str,most_resent_data:int):
    """
    Apply filters to results.
    """
    #filter section
    web_tools.click_element(locator=locators["section_button"])
    available_options = web_tools.get_text(locator=locators["section_options"]) 
    web_tools.click_element(locator=locators["section_target_option"].replace("index","1"))
    web_tools.click_element(locator=locators["section_button"])
    
    #filter date
    web_tools.click_element(locator=locators["date_range_options_list"])
    web_tools.click_element(locator=locators["specific_date_option"])
    prior_date = today_date - relativedelta(months=most_resent_data)
    web_tools.input_text(locator=locators["from_date_result"],text=prior_date.strftime("%m/%d/%Y")) 
    web_tools.input_text(locator=locators["to_date_result"],text=today_date.strftime("%m/%d/%Y")) 
    
    #sort form newes to older
    web_tools.click_element(locator=locators["order_by_list"])
    web_tools.select_from_list_by_value(locators["order_by_list"],*["newest"])
    web_tools.click_element(locator=locators["order_by_list"])
    

def download_image(save_path:str,image_url):
    """
    request an image and save it
    """
    response = requests.get(image_url)
    if response.status_code == 200:
        # Save the image to the specified path
        with open(save_path, "wb") as f:
            f.write(response.content)
    else:
        built_in_tools.log_to_console(f"Failed to download the image ({image_url}).")
    
def save_data_to_excel(items:list):
    """
    save results in an excel file
    """
    if len(items)==0:
        return
    target_path = f"{'output'}/{today_date.strftime('%d_%m_%Y__%H')}.xlsx"
    
    table  = table_tools.create_table(columns=list(items[0].keys()),data=items)
    
        
    try:
        excel_tools.create_workbook(target_path,"xlsx")
    except:
        excel_tools.open_workbook(target_path,"xlsx")
    excel_tools.append_rows_to_worksheet(table)
    excel_tools.save_workbook()
    excel_tools.close_workbook()
def compress_results():
    shutil.make_archive(media_folder,"zip",media_folder)
    shutil.move(f"{media_folder}.zip",f".\output")
@task
def minimal_task():
    try:
        go_to_page()
    except:
        built_in_tools.log_to_console(f"{error_message} could not got to {search_page}, verify internet conection or page status")
        return

    work_items_tools.for_each_input_work_item(get_results,items_limit=5)

 
