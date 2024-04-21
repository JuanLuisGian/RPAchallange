from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from Resources.Variables import *
from RPA.Robocorp.utils import BuiltIn
from RPA.Robocorp.WorkItems import WorkItems , State
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from RPA.Excel.Files import Files
from RPA.Tables import Tables
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
money_patterns = [
    r'$\d{0,3}.?\d+', r'$\d{0,3}\,\d{0,3}.\d{0,3}', r'\d+\bUSD\b', r'\d+\dollars\b']
media_folder = ".\Images"
try:
    os.mkdir(media_folder)
except:
    pass

with open("Locators\\all_locators.json", 'r') as file:
    locators = json.load(file)

def go_to_page():
    """
    Goes to home page, where the robot can start
    it's search.
    """
    try:
        web_tools.open_available_browser(search_page)
    except Exception as ex:
        raise(Exception(f"Error- Page is not avaialbe, las error was{str(ex)}"))

def get_results():
    """
    Process all work items
    """
    try:
        go_to_page()
        search_params = work_items_tools.get_work_item_payload()
        built_in_tools.log_to_console(f"Searching data on parameters:\n{search_params}")

        found = search_item(search_topic=search_params["search_phrase"], 
            result_category=search_params["result_category"], 
            most_resent_data=int(search_params["oldest_month"]))
        
        if (found):
            filter_results(search_params["result_category"])
            fetch_more_result = scann_results(search_params["search_phrase"],int(search_params["oldest_month"]))
            while (fetch_more_result):
                #try_up to three time
                for atempt in range(max_atemps_to_scan_results):
                    try:
                        fetch_more_result = scann_results(search_params["search_phrase"],
                            int(search_params["oldest_month"]))
                        break
                    except:
                        built_in_tools.log_to_console(f"failed to scan results from page, attemp {atempt+1}/{max_atemps_to_scan_results}")
                        pass
                try:
                    web_tools.wait_until_element_is_enabled(locators["view_older_results_button"],15)
                    web_tools.click_element(locators["view_older_results_button"])
                except:
                    fetch_more_result = False

        work_items_tools.release_input_work_item(state=State.DONE,message="Data Scraped")
    except Exception as ex:
        built_in_tools.log_to_console("-An error has ocurred, please refer to the following error")
        built_in_tools.log_to_console(f"{str(ex)}")
        work_items_tools.release_input_work_item(state=State.FAILED,message=str(ex),exception_type="APPLICATION")

def search_item(search_topic: str, result_category: str, most_resent_data: int):
    """
    Searchs and detects if a results exist upon search
    search_topic: value to search for
    result_category: category to filter results
    most_resent_data: how old in terms of months th result can be(1,2,3 months ago)
    """
    # check input  validity
    if len(search_topic) == 0 or len(result_category) == 0 or most_resent_data > max_months:
        built_in_tools.log_to_console(
            f"Data not valid, Topic :{search_topic},category : {result_category}, months ago {most_resent_data}")
        built_in_tools.log_to_console(
            f"Max allowed moths{most_resent_data}"
        )
        return False
    # type result in page
    web_tools.click_element(locator=locators["search_icon"])
    web_tools.input_text(
        locator=locators["search_box"], text=search_topic, clear=True)
    web_tools.click_element(locator=locators["search_button"])
    try:
        if "No search results match the term " in web_tools.get_text(locators["No_availableresult"]):
            return False
    except:
        pass
    return True
    
def scann_results(search_phrase:str,months_ago:int):
    """
    Reads all resulting rows and extract metadata

    months_ago : how old (in months) the results can be
    search_phrase: value used to idefntiy the set on the excel
    """
    #wait for the table to load
    web_tools.wait_until_element_is_visible(
        locator=locators["available_items_result"],timeout=max_time_to_load_results)
    


    # scan up to max number of rows
    no_of_available_rows = web_tools.get_element_count(
        locator=locators["available_items_result"])
    results = [
    #A separator for every payload
         {"Title": "******", "Date": "*******", "Description": "New Search Begins",
                  "Picture": f"for tearm'{search_phrase}'", "PhrasesCount": "*******", "MoneyAmountPresent": False}
    ]

    # we must continue fetching row as far they complie with the rules
    continue_to_fetch_more_results = True

    for index in range(1, no_of_available_rows+1):
        result = {"Title": "", "Date": "", "Description": "",
                  "Picture": "", "PhrasesCount": "", "MoneyAmountPresent": False}

        result["Title"] = web_tools.get_text(
            locator=locators["title_result_target"].replace("index", str(index)))
        
        #Download the image bia URL 
        try:
            web_tools.wait_until_element_is_visible(
                locator=locators["image_result_targe"].replace("index", 
                str(index)),timeout=5)
            
            image_download_url = web_tools.get_element_attribute(
                locator=locators["image_result_targe"].replace("index", str(index)), attribute="src")
            
            result["Picture"] = image_download_url.split("/")[-1]
            
            download_image(
                f"{media_folder}/{result['Picture']}", image_download_url)
        except Exception as ex:
            built_in_tools.log_to_console(f"Unable to dowload an image, last exception:\n {str(ex)}")
            result["Picture"] = "Not available"

        result["Description"] = "Not available"

        result["PhrasesCount"] = len(result["Title"].split(" "))

        result["Date"] = web_tools.get_text(
            locator=locators["date_result_target"].replace("index", str(index)))
        
        try:
            #expected formant, Month day, year (Ex. September 16, 2023)
            result["Date"] = datetime.strptime(result["Date"], "%B %d, %Y")
        except:
            result["Date"] = today_date

        
        for pattern in money_patterns:
            result["MoneyAmountPresent"] = len(
                re.findall(pattern, result["Description"])) > 0
            if (result["MoneyAmountPresent"]):
                break
        #check of date is range
        if not relativedelta(today_date, result["Date"]).months <= months_ago:
            continue_to_fetch_more_results = False

        results.append(result)
    save_data_to_excel(results)
    return continue_to_fetch_more_results

def filter_results(result_category: str):
    """
    Apply filters to results.
    result_category: category to filter the results on

    the bot tries to get only the newest data
    """
    try:
        # filter section
        web_tools.click_element(locator=locators["news_section_filter_button"])
        available_options = list(web_tools.get_text(locator=locators["news_section_filter_list"]).split("\n"))
        target_index = 0
        for option in available_options:
            if result_category in option:
                target_index =str(available_options.index(option)+1)
                break
        if target_index:
            #if an option is selected, the options list will close automaticly
            web_tools.click_element(
            locator=locators["news_section_filter_option"].replace("index", target_index))
        else:
            #if no option is selected, the options list mus be close
            web_tools.click_element(
            locator=locators["news_section_filter_option"].replace("index", "1"))


        # sort form newes to older
        web_tools.click_element(locator=locators["sort_by_button"])
        available_options = list(web_tools.get_text(locator=locators["sort_by_options_list"]).split("\n"))
        target_index = 0
        for option in available_options:
            if "Newest" in option:
                target_index =str(available_options.index(option)+1)
                break
        if target_index:
            #if an option is selected, the options list will close automaticly
            web_tools.click_element(
            locator=locators["sort_by_options_option"].replace("index", target_index))
        else:
            #if no option is selected, the options list mus be close
            web_tools.click_element(
            locator=locators["news_section_filter_option"].replace("index", "1"))
    except:
        built_in_tools.log_to_console("there was an error while trying to sort by date or filter section")

def download_image(save_path: str, image_url):
    """
    request an image and save it
    """
    response = requests.get(image_url,timeout=max_time_download_image)
    if response.status_code == 200:
        # Save the image to the specified path
        with open(save_path, "wb") as f:
            f.write(response.content)
    else:
        built_in_tools.log_to_console(
            f"Failed to download the image ({image_url}).")

def save_data_to_excel(items: list):
    """
    save results in an excel file
    """
    if len(items) == 0:
        return
    target_path = f"{'output'}/{today_date.strftime('%d_%m_%Y__%H')}.xlsx"

    table = table_tools.create_table(columns=list(items[0].keys()), data=items)

    try:
        excel_tools.open_workbook(target_path, "xlsx")
    except:
        excel_tools.create_workbook(target_path, "xlsx")
    excel_tools.append_rows_to_worksheet(content=table,header=True)
    excel_tools.save_workbook()
    excel_tools.close_workbook()

def compress_results():
    shutil.make_archive(media_folder, "zip", media_folder)
    shutil.move(f"{media_folder}.zip", f".\output")
    #clear media file afther zip
    shutil.rmtree(media_folder)
@task
def minimal_task():

    work_items_tools.for_each_input_work_item(get_results, items_limit=5)