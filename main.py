import pandas as pd
import requests
import json
from datetime import datetime
from module.To_Excel import To_xlsx
from module.Data_processing import Dataframe_list
from module.Share_file import FileSharer 


if __name__ == "__main__":  
    # 設定欲查詢的縣市
    query_list = ["台北市", "新北市", "台南市", "高雄市", "屏東縣", "台中市"]
    for i, County in enumerate(query_list):
        if "台" in County:
            County = County.replace('台', '臺')
        query_list[i] = County

    
    # 環保署PM2.5 API資料來源
    url = "https://data.epa.gov.tw/api/v1/aqx_p_432?limit=1000&api_key=9be7b239-557b-4c10-9775-78cadfc555e9&format=json"

    # 取得環保署PM2.5 API的回應
    res = requests.get(url)

    # 若res == 200表示成功
    if res.status_code == requests.codes.ok:
        print("下載資料成功")

    # 將下載之資料轉成python物件，並只截取需要的series以縮減資料量 (減少記憶體使用)
    d = res.text
    o_data = json.loads(d)
    o_data = o_data["records"]
    
    
    #o_data = pd.read_csv("aqi_data.csv")
    
    o_data = pd.DataFrame(o_data)
    data = o_data[["County", "SiteName", "PublishTime", "PM2.5", "AQI", "Status"]]

    # Rename columns 
    data.columns = ["County", "SiteName", "Date:Time", "PM2.5(unit: μg/m3)", "AQI", "Status"]

    # 填補缺失值
    for column in data.columns:
        for row_index, d in enumerate(data[column]):
            if d == "ND": #csv
                data.loc[row_index, column] = 0
            elif d == "": # json
                data.loc[row_index, column] = 0

    # Create a list to store datas we queried (each county's data will be stored as an element).
    county_df_list = []

    # 擷取資料中user查詢縣市之資料並印出查詢結果
    for county in query_list:
        county_data = data[data["County"] == county]
        county_df_list.append(county_data)
        print(county_data)

    # Construct a Dataframe_list class and use method "aqi_pm25_mean()" to calculate the mean of PM2.5 & AQI
    df_list = Dataframe_list(county_df_list)
    df_list.aqi_pm25_mean()  

    # Get today data + time (YYYY/MM/DD HH:MM:SS)
    date = datetime.today()
    date = date.date()
    
    # 轉換成xlsx格式並儲存
    tmp = To_xlsx(df_list.list, f'AQI_report_{date}.xlsx', query_list)
    tmp.write_to_xlsx()

    # Upload the report to filestack and let user can download it with the url
    filename = f'AQI_report_{date}.xlsx'
    file_sharer = FileSharer(filepath = rf"{filename}", api_key = 'A1Ha6kxtcTXOzUHD4eFUjz')
    print("Download pdf report by below url.")
    print(file_sharer.share())


    