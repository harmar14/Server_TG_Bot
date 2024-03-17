# -*- coding: cp1251 -*-

import telebot
import wmi
import os
import subprocess
import pythoncom
from datetime import datetime

BOT_TOKEN = "<your_tg_bot_token>"

ALLOWED_USERS = ["<tg_user_id_1>", "<tg_user_id_2>", "tg_user_id_3"]

ENGINE_PROCESS = "<engine_process_name>"
SCHEDULER_PROCESS = "<scheduler_process_name>"

START_ENGINE_CMD = f"net.exe start {ENGINE_PROCESS}"
STOP_ENGINE_CMD = f"net.exe stop {ENGINE_PROCESS}"
RESTART_ENGINE_CMD = f"{STOP_ENGINE_CMD} && {START_ENGINE_CMD}"
START_SCHEDULER_CMD = f"net.exe start {SCHEDULER_PROCESS}"
STOP_SCHEDULER_CMD = f"net.exe stop {SCHEDULER_PROCESS}"
RESTART_SCHEDULER_CMD = f"{STOP_SCHEDULER_CMD} && {START_SCHEDULER_CMD}"

SERVER = "<host_name>"

LOG_PATH = "<log_path>"

bot = telebot.TeleBot(BOT_TOKEN)

@bot.message_handler()
def text(message):
    chat_id = message.chat.id
    user_id = message.from_user.id
    print(user_id)
    command = telebot.util.extract_command(message.text)
    if ( f"{user_id}" in ALLOWED_USERS ):
        if ( command is not None ):
            if (command == "open_log_dir"):
                # Getting logs from log direcory
                if ( os.path.exists(LOG_PATH) ):
                    bot.send_message(chat_id, "Looking for files...")
                    log_dir_lookup(chat_id)
                    log_msg = bot.send_message(chat_id, "Input file name to get file")
                    bot.register_next_step_handler(log_msg, send_log_file)
                else:
                    bot.send_message(chat_id, f"Cannot find path {LOG_PATH}")
            elif (command == "check_engine"):
                if ( check_process(ENGINE_PROCESS) ):
                    bot.send_message(chat_id, f"Engine ({ENGINE_PROCESS}) is running")
                else:
                    bot.send_message(chat_id, f"Engine ({ENGINE_PROCESS}) is not running")
            elif (command == "check_threads"):
                # Checking if trere are active threads
                if ( check_process(ENGINE_PROCESS) ):
                    bot.send_message(chat_id, "It will take 2-4 mins...")
                    threads_cnt = check_win_logs()
                    if ( threads_cnt > 0 ):
                        bot.send_message(chat_id, f"There are active threads ({threads_cnt})")
                    else:
                        bot.send_message(chat_id, "There are no active threads")
                else:
                    bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is not running")
            elif (command == "start_engine"):
                # Starting Engine
                if ( check_process(ENGINE_PROCESS) ):
                    bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is already running")
                else:
                    # If process is not running, we can run it
                    subprocess.call(START_ENGINE_CMD, shell=True)
                    if ( check_process(ENGINE_PROCESS) ):
                        bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is started")
                    else:
                        bot.send_message(chat_id, f"Could not start process {ENGINE_PROCESS}")
            elif (command == "stop_engine"):
                # Stopping Engine
                if ( check_process(ENGINE_PROCESS) ):
                    # Is frocess is running, we can stop it
                    subprocess.call(STOP_ENGINE_CMD, shell=True)
                    if ( check_process(ENGINE_PROCESS) ):
                        bot.send_message(chat_id, f"Could not stop process {ENGINE_PROCESS}")
                    else:
                        bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is stopped")
                else:
                    bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is already stopped")
            elif (command == "restart_engine"):
                # Restarting Engine
                if ( check_process(ENGINE_PROCESS) ):
                    # If process is running, we can restart it
                    subprocess.call(RESTART_ENGINE_CMD, shell=True)
                    if ( check_process(ENGINE_PROCESS) ):
                        bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is restarted")
                    else:
                        bot.send_message(chat_id, f"Could not restart process {ENGINE_PROCESS}")
                else:
                    bot.send_message(chat_id, f"Process {ENGINE_PROCESS} is not running, try /start_engine")
            elif (command == "check_scheduler"):
                # Checking if Scheduler is running
                if ( check_process(SCHEDULER_PROCESS) ):
                    bot.send_message(chat_id, f"Scheduler ({SCHEDULER_PROCESS}) is running")
                else:
                    bot.send_message(chat_id, f"Scheduler ({SCHEDULER_PROCESS}) is not running")
            elif (command == "start_scheduler"):
                # Starting Scheduler
                if ( check_process(SCHEDULER_PROCESS) ):
                    bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is already running")
                else:
                    # If process is not running, we can run it
                    subprocess.call(START_SCHEDULER_CMD, shell=True)
                    if ( check_process(SCHEDULER_PROCESS) ):
                        bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is running")
                    else:
                        bot.send_message(chat_id, f"Could not start process {SCHEDULER_PROCESS}")
            elif (command == "stop_scheduler"):
                # Stopping Scheduler
                if ( check_process(SCHEDULER_PROCESS) ):
                    # If process is running, we can stop it
                    subprocess.call(STOP_SCHEDULER_CMD, shell=True)
                    if ( check_process(SCHEDULER_PROCESS) ):
                        bot.send_message(chat_id, f"Could not stop process {SCHEDULER_PROCESS}")
                    else:
                        bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is stopped")
                else:
                    bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is already stopped")
            elif (command == "restart_scheduler"):
                # Restarting Scheduler
                if ( check_process(SCHEDULER_PROCESS) ):
                    # If process is running, we can restart it
                    subprocess.call(RESTART_SCHEDULER_CMD, shell=True)
                    if ( check_scheduler() ):
                        bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is restarted")
                    else:
                        bot.send_message(chat_id, f"Could not restart process {SCHEDULER_PROCESS}")
                else:
                    bot.send_message(chat_id, f"Process {SCHEDULER_PROCESS} is not active, try /start_scheduler")
            else:
                bot.send_message(chat_id, "What does it mean? Please choose command from menu")
        else:
            bot.send_message(chat_id, "Waiting for a command...")
    else:
        bot.send_message(chat_id, "Mom taught me not to talk to strangers")

def log_dir_lookup(chat_id):
    # Send all log file names to chat one by one
    # (so that we will not have to deal with max message length)
    for file in os.listdir(LOG_PATH):
        time_modified = datetime.fromtimestamp(os.path.getmtime(f"{LOG_PATH}/{file}")).strftime("%m.%d.%Y %H:%M:%S")
        bot.send_message(chat_id, f"{file}\nModified: {time_modified}")

def send_log_file(message):
    # Send file to chat (is it is not too big)
    chat_id = message.chat.id
    file_name = message.text
    file_path = f"{LOG_PATH}/{file_name}"
    if ( os.path.exists(file_path) ):
        try:
            bot.send_document(chat_id, open(file_path,'rb'))
        except:
            bot.send_message(chat_id, "Exception. Looks like file is too big")
    else:
        bot.send_message(chat_id, "Wrong file name. Please try again")

def check_process(process):
    # Check if service is running or not
    pythoncom.CoInitialize()
    connection = wmi.WMI(SERVER)
    
    wql = (f"SELECT * FROM Win32_Process WHERE Name='{process}.exe'")
    query_res = connection.query(wql)

    if ( len(query_res) > 0 ):
        return True
    else:
        return False

def check_win_logs():
    # Get number of active threads from Windows Application logs
    ended_threads = []
    active_threads = []
    
    pythoncom.CoInitialize()
    connection = wmi.WMI(SERVER)
    
    wql = ("SELECT Message FROM Win32_NTLogEvent WHERE Logfile='Application' AND EventIdentifier='1'")
    query_res = connection.query(wql)
    
    # You should check logs in Windows Event Viewer
    # to come up with parsing algorithm.
    # Here is what workes for me:
    
    for item in query_res:
        msg = item.Message
        print(msg)
        
        if ( msg == "Service started." ):
            if ( len(active_threads) > 0 ):
                return len(active_threads)
            else:
                return 0
        
        msg_words = msg.split(' ')
        if ( msg_words[0] == "Thread" ):
            if ( msg_words[2] == "started." ):
                # Log about starting the process
                if ( msg_words[1] in ended_threads ):
                    # Ended thread
                    ended_threads.remove(msg_words[1])
                else:
                    # Active thread
                    active_threads.append(msg_words[1])
            elif ( msg_words[2] == "ended." ):
                # Log about ending the process
                ended_threads.append(msg_words[1])
    
    if ( len(active_threads) > 0 ):
        return len(active_threads)
    else:
        return 0

bot.polling(none_stop=True)
