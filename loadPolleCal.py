import pandas as pd
from icalendar import Calendar, Event
import os

# function to map index to new index
def map_index(index):
        # keep the indices which start with "Woche" or "KW" as they are
        if (str(index).startswith('Woche')) or (str(index).startswith('KW')):
            return index
        else:
            if sheet_name in ["Februar", "März", "April"]:
                # transform index "Vormittags/Mittags" to timestamp of 12:00, "Mittags/Nachmittags" to timestamp 14:00, "Nachmittags/Abends" to timestamp of 19:00, "Abends/Nachts" to timestamp of 22:00, "Nachts/Vormittags" to timestamp of 9:00 and 
                return {'Vormittags/Mittags': '12:00', 'Mittags/Nachmittags': '14:00', 'Nachmittags/Abends': '19:00', 'Abends/Nachts': '22:00', 'Nachts/Vormittags nächste Tag': '9:00', 'Nachts/Vormittas nächster Tag': '9:00'}.get(index, index)
            else:
                return {'Vormittags': '09:00', 'Mittags': '12:00', 'Nachmittags': '14:00', 'Abends': '18:00', 'Nachts': '22:00'}.get(index, index)

# function to return new week df
def return_newWeek_df(df):
        # make df of first 5 rows and remove from original dataframe
        df1 = df.iloc[:5]
        df = df.iloc[5:]
        #print(len(df))
        if len(df) == 0:
            return df1, df
        # make first row which starts with "Woche" to column names
        df.columns = df.iloc[0]
        df = df.drop(df.index[0])
        return df1, df


################# Beginning of script ######################
save_to_Ical = True

# import excel file
xl = pd.ExcelFile('/Users/xxxx/Downloads/Terminkalender.xlsx')

sheet_names = ["Mai"]# xl.sheet_names  # get all sheet names ####["September"] #
# run through all sheets of file
for sheet_name in sheet_names:
    # import sheet of file
    df = pd.read_excel('/Users/xxxx/Downloads/Terminkalender.xlsx', sheet_name=sheet_name)
   
    # drop columns which start with "Pauline"
    df = df[df.columns.drop(list(df.filter(regex='Pauline')))]

    # drop rows without index and values
    df = df.dropna(how='all', axis=0)

    # first row as column names and drop it
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    
    first_col = df.columns[0] # save the first column name
    # drop rows with entries "Notiz", "Spooooortchallenge", "Wetter" in first column
    df = df.drop(df[df[first_col].isin(["Notiz", "Spooooortchallenge", "Wetter"])].index)
    # make first column to index
    df = df.set_index(first_col)
    # map index to new index
    df.index = df.index.map(map_index)

    # build a df for each week and save df in list
    nr_weeks = len(df) // 5  # 5 rows (times) per week
    weeks = []
    for i in range(nr_weeks):
        df1, df = return_newWeek_df(df)
        weeks.append(df1)

    events_week = []
    for week in weeks:
        # add to each index the string ":00" to the end of the string
        week.index.name = None
        week.index = pd.to_timedelta(week.index + ":00")

        # remove columns which are NaT or NaN
        week = week.dropna(axis=1, how='all')
        # create new df with columns "Date" and "Event"
        df_week1 = pd.DataFrame(columns=["Date","Event"])
        # run through columns of week df, add index time to column date and save as "Date" of that row, then save to df_week1
        for column in week.columns:
            # if "mir dir" in string "week[column]" replace with "mit Pauline"
            week[column] = week[column].str.replace("mir dir", "mit Pauline")
            # add index time to column date and save as "Date" of that row, save "Event"
            df_week_temp = pd.DataFrame({"Date": pd.to_datetime(column) + week.index, "Event": week[column]})

            if sheet_name in ["Februar", "März", "April"]:
                # if time in "Date" is 09:00:00, then add one day to the date
                df_week_temp.loc[df_week_temp["Date"].dt.time == pd.to_datetime("09:00:00").time(), "Date"] += pd.DateOffset(days=1)

            # concatinate df_week_temp to df_week1
            df_week1 = pd.concat([df_week1, df_week_temp])

        # reset index of df_week1
        df_week1 = df_week1.reset_index(drop=True)
        # remove rows where Event is NaN
        df_week1 = df_week1.dropna(subset=["Event"])
        events_week.append(df_week1)

    ##################### Save to Ical ############################
    if save_to_Ical == True:
        cal = Calendar()
        cal.add('prodid', sheet_name)
        cal.add('version', '2.0')

        # run through the list of weeks
        for df_week1 in events_week:
            # run through each row and take the date and event and save to calendar
            for index, row in df_week1.iterrows():
                event = Event()
                event.add('summary', row["Event"])
                event.add('dtstart', row["Date"])
                event.add('dtend', row["Date"] + pd.DateOffset(hours=2))
                cal.add_component(event)

                # write to disk
                directory = "/events"    
                event_name = sheet_name
            
                file_path = os.path.join(directory, event_name + '.ics')
                if not os.path.exists(file_path):
                    # create file

                    with open(event_name + '.ics', 'wb') as f:
                        # Perform any necessary operations on the file
                        f.write(cal.to_ical())
                        pass

                #print(f"File saved: {file_path}")
                
                f.close()
        #print(cal.to_ical().decode("utf-8")) 
        print("############## Sheet "+ sheet_name +" done #################")
