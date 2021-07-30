import pandas as pd

class To_xlsx:

    '''
    df_list中的元素為dataframe (一個縣市一個dataframe)，寫入excel時各個dataframe分別寫入AQI report檔案中不同的sheet
    '''

    def __init__(self, df_list, file_name, city_list):
        self.df_list = df_list
        self.file_name = file_name
        self.city_list = city_list

    def write_to_xlsx(self):
        #Create a Pandas Excel writer using XlsxWriter as the engine.
        output_path = rf"/Users/wangyaojun/Desktop/{self.file_name}"
      
        writer = pd.ExcelWriter(output_path, engine = 'xlsxwriter')
        df_list = []

        
        for df in self.df_list:
            #df[["PM2.5(unit: μg/m3)", "AQI"]] = df[["PM2.5(unit: μg/m3)", "AQI"]].astype(int)
            county = df.iat[0, 0]

            # Convert the dataframe to an XlsxWriter Excel object.
            # Notice: Default header has been turned off and skip one row to allow us to insert a user defined header
            df.to_excel(writer, sheet_name = county, startrow = 1, header = False, index = False)

            # Get the xlsxwriter workbook and worksheet objects.
            workbook  = writer.book
            worksheet = writer.sheets[county]

            # Header format
            header_format = workbook.add_format({
                                                'font_size': 20,
                                                'bold': True,
                                                'text_wrap': True,
                                                'align': 'center',
                                                'valign': 'vcenter',
                                                'fg_color': '#D7E4BC',
                                                'border': 1
                                                })

            # Data format
            data_format = workbook.add_format({
                                                'font_size': 18,
                                                'bold': True,
                                                'text_wrap': True,
                                                'align': 'center',
                                                'valign': 'vcenter',
                                                'border': 1
                                            })
            
            # Set rows format
            for i in range(len(df)+1):
                if i == 0:
                    # Set label row format
                    worksheet.set_row(i, height = 60, cell_format = header_format)
                else:
                    worksheet.set_row(i, height = 35, cell_format = data_format)

            column_width = [20, 20, 40, 40, 20, 45]
            # Set columns width
            for i, j in enumerate(column_width):
                worksheet.set_column(i, i, width = j)
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format) 

            # Python 色碼請參考: https://blog.csdn.net/guduruyu/article/details/77836173
            # 橘色字體
            cell_format = workbook.add_format({'font_color': "#FFA500"})
            # 暗紅色字體
            cell_format2 = workbook.add_format({'font_color': "#930000"})
            # 紫色字體
            cell_format3 = workbook.add_format({'font_color': "#800080"})
            # 紅色字體
            cell_format4 = workbook.add_format({'font_color': "#FF0000"})
            

            # conditional_format(first_row, first_col, last_row, last_col, options)  
            # 若AQI > 100 以上將AQI欄位數值以橘色顯示 (對敏感族群不健康)                               
            worksheet.conditional_format(1, 4, len(df)+1, 4, {"type": "cell", "criteria": "between", "minimum": 101, "maximum": 150, 'format': cell_format})
            # 若AQI > 150 以上將AQI欄位數值以暗紅色顯示 (對所有族群不健康)
            worksheet.conditional_format(1, 4, len(df)+1, 4, {"type": "cell", "criteria": "between", "minimum": 151, "maximum": 200, 'format': cell_format2})
            # 若AQI > 200 以上將AQI欄位數值以紫色顯示 (非常不健康)
            worksheet.conditional_format(1, 4, len(df)+1, 4, {"type": "cell", "criteria": "between", "minimum": 201, "maximum": 300, 'format': cell_format3})
            # 若AQI > 300 以上將AQI欄位數值以紅色顯示 (有害)
            worksheet.conditional_format(1, 4, len(df)+1, 4, {"type": "cell", "criteria": ">", "value": 301, 'format': cell_format4})

            # Status欄位字體顏色與AQI欄位匹配 
            # 橘色: 對敏感族群不健康
            worksheet.conditional_format(1, 5, len(df)+1, 5, {"type": "text", "criteria": "containing", "value": "敏感", 'format': cell_format})
            # 暗紅色: 對所有族群不健康
            worksheet.conditional_format(1, 5, len(df)+1, 5, {"type": "text", "criteria": "containing", "value": "所有", 'format': cell_format2})
            # 紫色: 非常不健康
            worksheet.conditional_format(1, 5, len(df)+1, 5, {"type": "text", "criteria": "containing", "value": "非常", 'format': cell_format3})
            # 紅色: 有害
            worksheet.conditional_format(1, 5, len(df)+1, 5, {"type": "text", "criteria": "containing", "value": "有害", 'format': cell_format4})

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
