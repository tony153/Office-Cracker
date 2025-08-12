import json
import os
import sys
from mainfunction import *
import inspect
import traceback

command = None


def resource_path(relative_path):
    """取得打包後的資源路徑"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包後的暫存目錄
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        # 開發階段直接使用原始路徑
        return os.path.join(os.path.abspath("."), relative_path)
def load_command():
    global command
    command_file = resource_path("command.json")
    with open(command_file,"r", encoding="UTF-8") as command_file:
        command_str = command_file.read()
        command = json.loads(command_str)
        print("[✔] load commands done.")

def run_function(function_name, userinput):
    arg_spec = inspect.signature(globals()[function_name])
    if len(arg_spec.parameters.keys()) > 0:
        globals()[function_name](userinput)
    else:
        globals()[function_name]()

def print_prompt(prompts):
    result = ""
    for index,prompts in enumerate(prompts):
        if index == 0:
            result = prompts
        else:
            result += "/"+prompts
    return result+">"

def print_help(current_command):
    command_counter = 0
    for key in current_command.keys():
        if type(current_command[key]) is dict:
            command_counter +=1
            print("💠"+key+":")
            if "_description" in current_command[key].keys():
                print("\t"+current_command[key]["_description"])
            else:
                print("\t! 指令沒有說明")

    if command_counter == 0:
        print("\t!不存在子指令 或這是執行類指令")
    else:
        print("💠?:\t指令查詢")
        print("💠q:\t返回上一層指令/(第一層)為結束程序")

def quit(current_command, command_path,prompt):
    if len(command_path) != 0:
        command_path.pop()
        prompt.pop()
        for comma in command_path:
            current_command = command[comma]
        if len(command_path) == 0:
            current_command = command
    else:
        exit()
    return current_command, command_path, prompt

def match_command(userinput,current_command):
    match_command_list = []
    for available_command in current_command.keys():
        if available_command.startswith(userinput):
            match_command_list.append(available_command)

    if len(match_command_list) == 1:
        return match_command_list[0]
    elif len(match_command_list) > 1:
        print("\t",end="存在以下指令\n\t")
        print("    ".join(match_command_list))
        return False
    else:
        return False

    #print(current_command.keys())

def sys_clear():
    os.system("cls")
def start():
    load_command()
    if "init" in globals():
        init()

    prompt = [command["_appname"]]
    command_path = []
    current_command = command

    while(True):
        user_input = input(print_prompt(prompt))
        if user_input == "":
            continue
        elif user_input == "?":
            print_help(current_command)
        elif user_input == "q":
            current_command, command_path, prompt= quit(current_command, command_path, prompt)
        elif user_input == "cls":
            sys_clear()
        elif user_input.startswith("_") and user_input.split(" ")[0] in current_command.keys():
            command_str = user_input.split(" ")[0]
            print(current_command[command_str])
        elif user_input.startswith("_"):
            print("\t! 參數不存在")
        elif match_command(user_input.split(" ")[0], current_command):
            user_input_command = match_command(user_input.split(" ")[0], current_command)
            current_command = current_command[user_input_command]
            command_path.append(user_input_command)
            prompt.append(current_command["_name"])
            if "_function" in current_command.keys():
                try:
                    run_function(current_command["_function"],user_input)
                    current_command, command_path, prompt = quit(current_command,command_path,prompt)
                except Exception as e:
                    traceback.print_exc()
                    current_command, command_path, prompt = quit(current_command, command_path, prompt)
        else:
            print("\t! 指令不存在/指令不明碓")

