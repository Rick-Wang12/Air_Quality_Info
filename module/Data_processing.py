import pandas as pd

class Dataframe_list:

    def __init__(self, df_list):
        self.list = df_list

    def aqi_pm25_mean(self):
        # 關閉警告
        # pd.options.mode.chained_assignment = None
        for i, d in enumerate(self.list):

            # 將PM2.5與AQI數值轉換為整數型態
            d[["PM2.5(unit: μg/m3)","AQI"]] = d[["PM2.5(unit: μg/m3)","AQI"]].astype(int)

            # 計算PM2.5與AQI數值的平均數(四捨五入至小數點後第二位)
            tmp_mean = d[["PM2.5(unit: μg/m3)","AQI"]].mean().round(2)

            # 將tmp_mean存回dataframe中，在county column以Mean為row name
            tmp_mean["County"] = "各測站平均"
            #print(type(tmp_mean))
            #print(tmp_mean)

            d = d.append(tmp_mean, ignore_index = True)
            self.list[i] = d
    
        
            