import xlrd, os, re, sys
from time import sleep
from Tkinter import Tk
from tkFileDialog import askopenfilename, asksaveasfile, asksaveasfilename


Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing

xlsfile = askopenfilename()
book = xlrd.open_workbook(xlsfile, 'rb')

print "Opening the file ..."

jil_file = asksaveasfile(mode='wb', defaultextension=".jil")

sheet = book.sheet_by_index(0)



def auto_restart(row):
#auto_restart

    auto_restart_tuple =("auto_restart", "AUTO_RESTART", "Auto Restart", "Auto Restart", "auto restart", "AUTO RESTART")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in auto_restart_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A" and sheet.cell_value(row, columns) != 0 and sheet.cell_value(row, columns) != "0":

                    jil_file.write("auto_restart: "  + unicode(int(sheet.cell_value(row,columns))) + "\n")


                else:

                    continue

def timezone(row):
#timezone

    timezone_tuple =("timezone", "TIMEZONE", "Timezone", "Time Zone", "time zone", "time_zone", "Time_Zone", "TIME_ZONE", "TIME ZONE")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in timezone_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("timezone: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def priority(row):
#priority

    priority_tuple =("priority", "PRIORITY", "Priority")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in priority_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("priority: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def job_load(row):
#job_load

    job_load_tuple =("job_load", "JOB_LOAD", "Job_Load")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in job_load_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("job_load: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def heart_beat_interval(row):
#heart_beat_interval

    heart_beat_interval_tuple =("heart_beat_interval", "HEART_BEAT_INTERVAL", "Heart_Beat_Interval")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in heart_beat_interval_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("heart_beat_interval: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def profile(row):
#profile

    profile_tuple =("profile", "PROFILE", "Profile")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in profile_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("profile: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def job_terminator(row):
#job_terminator

    job_terminator_tuple =("job_terminator", "JOB_TERMINATOR", "Job_Terminator")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in job_terminator_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("job_terminator: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def box_terminator(row):
#box_terminator

    box_terminator_tuple =("box_terminator", "BOX_TERMINATOR", "Box_Terminator")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in box_terminator_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("box_terminator: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def term_run_time(row):
#term_run_time

    term_run_time_tuple =("term_run_time", "TERM_RUN_TIME", "Term_Run_Time")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in term_run_time_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("term_run_time: "  + unicode(int(sheet.cell_value(row,columns))) + "\n")


                else:

                    continue

def max_run_alarm(row):
#max_run_alarm

    max_run_alarm_tuple =("max_run_alarm", "MAX_RUN_ALARM", "Max_Run_Alarm", "MAX_RUNTIME_ALARM", "Max_Runtime_Alarm")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in max_run_alarm_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A" and filter(unicode.isdigit,unicode(sheet.cell_value(row,columns))) != "00":

                    number = filter(unicode.isdigit,unicode(sheet.cell_value(row,columns)))

                    if len(number) == 2:

                        jil_file.write("max_run_alarm: "  + number + "\n")

                    else:

                        jil_file.write("max_run_alarm: "  + number[:-1] + "\n")


                else:

                    continue

def min_run_alarm(row):
#min_run_alarm

    min_run_alarm_tuple =("min_run_alarm", "MIN_RUN_ALARM", "Min_Run_Alarm")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in min_run_alarm_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A" and sheet.cell_value(row, columns) != "0" and sheet.cell_value(row, columns) != 0:

                    jil_file.write("min_run_alarm: "  + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def alarm_if_fail(row):
#alarm_if_fail

    alarm_if_fail_tuple = ("alarm_if_fail", "ALARM_IF_FAIL", "Alarm If Fail")
    true_table = ("TRUE", "true", "True")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in alarm_if_fail_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if unicode(sheet.cell_value(row, columns)) != "" and unicode(sheet.cell_value(row, columns)) != "n/a" and unicode(sheet.cell_value(row, columns)) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A" and sheet.cell_value(row, columns) != "0" and sheet.cell_value(row, columns) != 0:


                    if unicode(sheet.cell_value(row, columns)) in true_table or unicode(int(sheet.cell_value(row, columns))) == "1":


                        jil_file.write("alarm_if_fail: " + "1" + "\n")


                else:

                    continue

def run_window(row):
#run_window

    run_window_tuple =("run_window", "RUN_WINDOW", "Run_Window")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in run_window_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("run_window: " + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")


                else:

                    continue

def start_mins(row):
#start_mins
    global date_condition
    start_mins_tuple =("start_mins", "START_MINS", "Start_Mins", "start mins")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in start_mins_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

					try:

						time = xlrd.xldate_as_tuple(sheet.cell_value(row, columns), book.datemode)



						if len(unicode(time[3])) == 2:

							if len(unicode(time[4])) == 2:

								jil_file.write("start_mins: " + '"' + unicode(time[3]) + ":" + unicode(time[4]) + '"' + "\n")

								date_condition = True

							elif len(unicode(time[4])) == 1:

								jil_file.write("start_mins: " + '"' + unicode(time[3]) + ":0" + unicode(time[4]) + '"' + "\n")

								date_condition = True

						elif len(unicode(time[3])) == 1:

							if len(unicode(time[4])) == 2:

								jil_file.write("start_mins: " + '"' + "0" + unicode(time[3]) + ":" + unicode(time[4]) + '"' + "\n")

								date_condition = True

							elif len(unicode(time[4])) == 1:

								jil_file.write("start_mins: " + '"' + "0" + unicode(time[3]) + ":0" + unicode(time[4]) + '"' + "\n")

								date_condition = True

					except:

						 jil_file.write("start_mins: " + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")

                else:

                    continue

def start_times(row):
#start_times
    global date_condition
    start_times_tuple =("start_times", "START_TIMES", "Start_Times", "start times")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):
            unicode(sheet.cell_value(rows - 1, columns)).strip()
            if unicode(sheet.cell_value(rows - 1, columns)).lower() in start_times_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

					try:

						time = xlrd.xldate_as_tuple(sheet.cell_value(row, columns), book.datemode)



						if len(unicode(time[3])) == 2:

							if len(unicode(time[4])) == 2:

								jil_file.write("start_times: " + '"' + unicode(time[3]) + ":" + unicode(time[4]) + '"' + "\n")

								date_condition = True

							elif len(unicode(time[4])) == 1:

								jil_file.write("start_times: " + '"' + unicode(time[3]) + ":0" + unicode(time[4]) + '"' + "\n")

								date_condition = True

						elif len(unicode(time[3])) == 1:

							if len(unicode(time[4])) == 2:

								jil_file.write("start_times: " + '"' + "0" + unicode(time[3]) + ":" + unicode(time[4]) + '"' + "\n")

								date_condition = True

							elif len(unicode(time[4])) == 1:

								jil_file.write("start_times: " + '"' + "0" + unicode(time[3]) + ":0" + unicode(time[4]) + '"' + "\n")

								date_condition = True

					except:

						 jil_file.write("start_times: " + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")

                else:

                    continue

def exclude_calendar(row):
#exclude_calendar
    global date_condition
    exclude_calendar_tuple =("exclude_calendar", "EXCLUDE_CALENDAR", "Exclude_Calendar")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in exclude_calendar_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("exclude_calendar: " + unicode(sheet.cell_value(row,columns)) + "\n")

                    date_condition = True

                else:

                    continue

def run_calendar(row):
#run_calendar
    global date_condition
    run_calendar_tuple =("run_calendar", "RUN_CALENDAR", "Run_Calendar")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in run_calendar_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("run_calendar: " + unicode(sheet.cell_value(row,columns)) + "\n")

                    date_condition = True

                else:

                    continue

def days_of_week(row):
#days_of_week
    global date_condition
    days_of_week_tuple =("days_of_week", "DAYS_OF_WEEKS", "Days_Of_Week", "days of week")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in days_of_week_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("days_of_week: " + unicode(sheet.cell_value(row,columns)) + "\n")

                    date_condition = True

                else:

                    continue

def run_days(row):
#run_days

    run_days_tuple = ("run_days", "RUN_DAYS", "Run_Days", "Run_days", "Run Days", "run days")

    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in run_days_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if unicode(sheet.cell_value(row, columns)) != "" and unicode(sheet.cell_value(row, columns)) != "n/a" and unicode(sheet.cell_value(row, columns)) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":


                    jil_file.write("days_of_week: " + unicode(sheet.cell_value(row, columns)) + "\n")


                else:

                    continue

def run_window(row):
#run_window

    run_window_tuple = ("run_window", "RUN_WINDOW", "Run_Window", "Run_window", "Run Window")

    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in run_window_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if unicode(sheet.cell_value(row, columns)) != "" and unicode(sheet.cell_value(row, columns)) != "n/a" and unicode(sheet.cell_value(row, columns)) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":


                    jil_file.write("run_window: " + '"' + unicode(sheet.cell_value(row, columns)) +  '"' + "\n")


                else:

                    continue

def std_in_file(row):
#std_in_file

    std_in_file_tuple =("std_in_file", "STD_IN_FILE", "Std In File")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in std_in_file_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("std_in_file: " + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")


                else:

                    continue

def std_err_file(row):
#std_err_file

    std_err_file_tuple =("std_err_file", "STD_ERR_FILE", "Std Err File")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in std_err_file_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                   if  sheet.cell_value(row, columns).find('"') == 0:

                        jil_file.write("std_err_file: "  + unicode(sheet.cell_value(row,columns)) + "\n")

                   else:

                        jil_file.write("std_err_file: "  + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")


                else:

                    continue

def std_out_file(row):
#std_out_file

    std_out_file_tuple =("std_out_file", "STD_OUT_FILE", "Std Out File")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in std_out_file_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    if  sheet.cell_value(row, columns).find('"') == 0:

                        jil_file.write("std_out_file: "  + unicode(sheet.cell_value(row,columns)) + "\n")

                    else:

                        jil_file.write("std_out_file: "  + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")


                else:

                    continue

def watch_file(row):
#watch_file

    watch_file_tuple =("watch_file", "WATCH_FILE", "Watch_File", "Watch File")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in watch_file_tuple:
                unicode(sheet.cell_value(row, columns)).rstrip()
                #print "'" + unicode(sheet.cell_value(row, columns)) + "'"
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("watch_file: " +  '"' + unicode(sheet.cell_value(row,columns)) +  '"' + "\n")


                else:

                    continue

def watch_file_min_size(row):
#watch_file_min_size

    watch_file_min_size_tuple =("watch_file_min_size", "WATCH_FILE_MIN_SIZE", "Watch_File_Min_size", "Watch File Min Size")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in watch_file_min_size_tuple:
                unicode(sheet.cell_value(row, columns)).strip()

                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("watch_file_min_size: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def watch_interval(row):
#watch_interval

    watch_interval_tuple =("watch_interval", "WATCH_INTERVAL", "Watch_Interval", "Watch Interval")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in watch_interval_tuple:

                unicode(sheet.cell_value(row, columns)).strip()

                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("watch_interval: " + unicode(int(sheet.cell_value(row,columns))) + "\n")

                else:

                    continue

def description(row):
#description

    description_tuple =("description", "DESCRIPTION", "Description")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in description_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":


                    if  sheet.cell_value(row, columns).find('"') == 0:

                        jil_file.write("description: "  + unicode(sheet.cell_value(row,columns)) + "\n")

                    else:

                        jil_file.write("description: " + '"' + unicode(sheet.cell_value(row,columns)) + '"' + "\n")

                else:

                    continue

def date_conditions(row):
#date_conditions
    global date_condition
    date_conditions_tuple =("date_conditions", "DATE_CONDITIONS", "Date_Conditions")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in date_conditions_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    #jil_file.write("date_conditions: " + unicode(sheet.cell_value(row,columns)) + "\n")
                    date_condition = True

                else:

                    continue

def condition(row):
#condition

    condition_tuple =("condition", "CONDITION", "Condition")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in condition_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("condition: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def auto_hold(row):
#auto_hold

    auto_hold_tuple = ("auto_hold", "Auto_Hold", "Auto_hold", "AUTO_HOLD")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in auto_hold_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("auto_hold: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def n_retrys(row):
#n_retrys

    n_retrys_tuple =("n_retrys", "N_RETRYS", "N_retrys")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in n_retrys_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("n_retrys: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def permission(row):
#permission

    permission_tuple = ("permission", "Permission", "PERMISSION")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):



            if unicode(sheet.cell_value(rows - 1, columns)).lower() in permission_tuple:
                    unicode(sheet.cell_value(row, columns)).strip()
                    if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                        jil_file.write("permission: " + unicode(sheet.cell_value(row,columns)) + "\n")

                    else:

                        continue
            else:

                jil_file.write("permission: " + "gx,wx" + "\n")
                break
        break

def command(row):
#command

    command_tuple = ("command", "COMMAND", "Command")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in command_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    if  sheet.cell_value(row, columns).find('"') == 0:

                        jil_file.write("command: "  + (sheet.cell_value(row,columns)) + "\n")

                    else:

                        jil_file.write("command: "  + '"' + (sheet.cell_value(row,columns)) + '"' + "\n")

                else:

                    continue

def owner(row):
#owner

    owner_tuple = ("owner", "OWNER", "Owner")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in owner_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("owner: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def machine(row):
#machine

    machine_tuple = ("machine", "MACHINE", "Machine")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in machine_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("machine: " + unicode(sheet.cell_value(row,columns)).lower() + "\n")


                else:

                    continue


def box_name(row):
#box_name
    box_name_tuple = ("box", "Box", "BOX")

    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in box_name_tuple:
                unicode(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != 'n\\a' and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("box_name: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def job_type(row):
#job_type
    job_type_tuple = ("job_type","Job_type", "job type", "Job_Type")
    for columns in range(sheet.ncols):

        for rows in range (sheet.nrows):

            if unicode(sheet.cell_value(rows - 1, columns)).lower() in job_type_tuple:
                str(sheet.cell_value(row, columns)).strip()
                if sheet.cell_value(row, columns) != "" and sheet.cell_value(row, columns) != "n/a" and sheet.cell_value(row, columns) != "N/A" and sheet.cell_value(row, columns) != "n\\a" and sheet.cell_value(row, columns) != "N\\A":

                    jil_file.write("insert_job: " + insert_job + "  job_type: " + unicode(sheet.cell_value(row,columns)) + "\n")


                else:

                    continue

def xls2jil():
#write down the header
    global date_condition
    job_name_tuple = ("job_name", "Job_name", "JOB_NAME", "job name")
    for column in range(sheet.ncols):

        for row in range (sheet.nrows):

            if str(sheet.cell_value(row - 1, column)).lower() in job_name_tuple:

                while row < sheet.nrows:

                    str(sheet.cell_value(row, column)).strip()
                    if sheet.cell_value(row, column) != "" and sheet.cell_value(row, column) != "n/a" and sheet.cell_value(row, column) != "N/A" and sheet.cell_value(row, column) != "n\\a" and sheet.cell_value(row, column) != "N\\A":

                        jil_file.write("/* ----------------- " + sheet.cell_value(row,column) + " ----------------- */\n")
                        jil_file.write("\n")
                        global insert_job
                        insert_job = sheet.cell_value(row, column)



                        date_condition = False
                        job_type(row); box_name(row); owner(row); permission(row); machine(row); n_retrys(row); auto_hold(row);
                        command(row); condition(row); date_conditions(row); days_of_week(row);  run_calendar(row); exclude_calendar(row); start_times(row); start_mins(row);
                        run_window(row); run_days(row); description(row); term_run_time(row);
                        box_terminator(row); job_terminator(row); std_in_file(row); std_out_file(row); std_err_file(row); watch_file(row);
                        watch_file_min_size(row); watch_interval(row); min_run_alarm(row); max_run_alarm(row); alarm_if_fail(row);
                        profile(row); heart_beat_interval(row); job_load(row); priority(row); timezone(row); auto_restart(row);


                        if date_condition is True:

                            jil_file.write("date_conditions: 1\n")



                        jil_file.write("\n")
                        jil_file.write("\n")

                        row += 1


                    else:


                        row += 1
                        continue


xls2jil()

print "Closing the file ..."

print "\n"
print "Done!"
print "\n"

jil_file.close()

file = str(jil_file)

path = re.findall('\'(.+?)\'', file)

new_file = path[0]

os.startfile(new_file)

sleep(1.0)
