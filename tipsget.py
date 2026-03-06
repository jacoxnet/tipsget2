import requests
import pandas as pd
from datetime import datetime
import os
from dateutil.relativedelta import relativedelta
from calendar import monthrange

filename = "Test TIPS Book.xlsx"
default = {}
baseUrlt = 'https://api.fiscaldata.treasury.gov/services/api/fiscal_service'
summary_endpoint = '/v1/accounting/od/tips_cpi_data_summary'
details_endpoint = '/v1/accounting/od/tips_cpi_data_detail'
baseUrlfed = 'https://api.stlouisfed.org/fred/series/observations'
fedAPIkey = os.getenv("SFEDKEY")
ifields = ["index_date", "index_ratio"]
mikefields = {"cusip": "Cusip", "interest_rate": "Coupon", "maturity_date": "Maturity Date", 
              "security_term": "Term", "series": "Series", "original_issue_date": "Issue Date",
              "index_ratio": "Inflation Factor", "index_date": "Inflation Date", "ref_cpi_on_dated_date": "Dated date CPI-U"}
my_tips = []

# return date in string format YYYY-MM-DD for use in API calls and file output
def convert_date(the_date):
    return the_date.strftime('%Y-%m-%d')

def get_all_tips():
    print("Getting summary TIPS data...")
    API = baseUrlt + summary_endpoint
    response = requests.get(API)
    tips_list = response.json()["data"]
    return tips_list

# return list of indexes with index date and index ratio for the_date
def get_indexes(the_date):
    print("Getting index details ...")
    API = baseUrlt + details_endpoint
    params = {"filter": "index_date:eq:" + convert_date(the_date)}
    response = requests.get(API, params=params)
    index_list = response.json()["data"]
    return index_list

# return CPI-U for the_date which is used in calculating TIPS interest between index date and maturity date
def get_cpiu(the_date):
    print("Getting CPI-U data ...")
    # calculate date 3 months prior to the_date which is used in TIPS
    cpu_date = the_date - relativedelta(months=3)
    API = baseUrlfed
    params = {"series_id": "CPIAUCNS", "observation_start": convert_date(cpu_date), 
              "observation_end": convert_date(the_date), "api_key": fedAPIkey,
              "file_type": "json"}
    response = requests.get(API, params=params)
    cpiu_data = response.json()
    # calculate daily increase in CPI-U for use in calculating TIPS interest between index date and maturity date
    ob1 = float(cpiu_data["observations"][0]["value"])
    ob2 = float(cpiu_data["observations"][1]["value"])
    print("ob1: ", ob1, "ob2: ", ob2)
    daily_cpiu_inc = (ob2 - ob1) / monthrange(the_date.year, the_date.month)[1]
    print("daily_cpiu_inc: ", daily_cpiu_inc)
    # return daily increase times days so far in this month to get increase in CPI-U since index date
    return ob1 + daily_cpiu_inc * (the_date.day - 1)

# search index_list for cusip and return index info if found, otherwise return default
def find_index(cusip, index_list):
    for index_item in index_list:
        if index_item["cusip"] == cusip:
            return index_item
    return default

# write my_tips to DownloadedData sheet in xlsx file
def writefile(tips):
    print("Writing to DownloadedData sheet in xlsx file ...")
    df = pd.DataFrame(tips)
    with pd.ExcelWriter(filename, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='DownloadedData', index=False)

def main():
    idate = datetime.now()
    tips_list = get_all_tips()
    print ("Total tips received :", len(tips_list))
    # go thru recovered fields and place selected ones in my_tips
    for tip in tips_list:
        my_tip = {}
        for fieldname in mikefields.keys():
            if tip.get(fieldname):
                my_tip[mikefields[fieldname]] = tip[fieldname]
        my_tips.append(my_tip)
    index_list = get_indexes(idate)
    print ("total indexes received: ", len(index_list))
    # go thru tips and search for and recover index info
    cpiu = get_cpiu(idate)
    for tip in my_tips:
        index = find_index(tip[mikefields["cusip"]], index_list)
        tip[mikefields["index_ratio"]] = index.get("index_ratio")
        tip[mikefields["index_date"]] = index.get("index_date")
        if tip[mikefields["index_ratio"]]:
            tip["Adjusted Principal"] = int(float(tip[mikefields["index_ratio"]]) * 100000) / 100
        else:
            tip["Adjusted Principal"] = None
        tip["Current CPIU"] = cpiu
        tip["Calculated Inflation Factor"] = round((cpiu / float(tip["Dated date CPI-U"])), 5)
        # tip["Calculated Inflation Factor"] = cpiu / float(tip["Dated date CPI-U"])
    my_tips.sort(key=lambda x: x["Maturity Date"])
    writefile(my_tips)
    print("All done.")

if __name__ == "__main__":
    main()