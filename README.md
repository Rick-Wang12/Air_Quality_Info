
### AQI report

#### Description:
   * Download AQI data from the Environmental Protection Administration, and slice datas of county(s) we queried from it.

   * Columns are named as "County", "SiteName", "Date:Time", "PM2.5(unit: μg/m3)", "AQI", "Status"

#### Objects:
   * Data_processing:
        Calculate mean, summary, etc. for datas in Dataframes_list <br/>

   * To_Excel:
        Set the Excel format(font, width of cell, etc.) and write the dataframe into the excel file as .xlsx format.
        Stored each city AQI report as independent worksheet. Besides, we pass a list stored dataframes, which stored each county 
        we queried as a element. <br/> <br/>

        "write_to_xlsx": This method will output each dataframe as a independent worksheet in the excel file. <br/> <br/>
    
   * Share_file:
        Upload the file to internet and let user download the file by the url. <br/>

    
