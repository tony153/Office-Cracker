import json
import os.path
import subprocess

import getOfficeHashString
import officeCrackPW
import officeCrackPW_mp
import print_table
import datetime

config_filename = "config.json"
CommonPasswordsList = None
BasesWordsList = None
user_config = {
    "CommonPasswordsList_file": "-",
    "BasesWordsList_file": "-",
    "BruteForceLengthStart": 3,
    "BruteForceLengthEnd": 6,
}


def init():
    load_option()
    print_banner(title="Welcome to Office Cracker", version="v1.0", author="~", show_time=False)

def print_banner(
        title="Office Cracker",
        version="v1.0",
        author="",
        show_time=True
):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") if show_time else ""

    banner_lines = [
        "┌" + "─" * 50 + "┐",
        f"│  {title.center(46)}  │",
        f"│  {'Version: ' + version:<46}  │",
        f"│  {'Author: ' + author:<46}  │",
    ]

    if show_time:
        banner_lines.append(f"│  {'Time: ' + timestamp:<36}  │")

    banner_lines.append("└" + "─" * 50 + "┘")

    print("\n".join(banner_lines))


def prompt(msg):
    value = input(msg+"(-q:退出)>")
    return None if value == "-q" else value


def get_Office_Hash_String():
    file_path = prompt("檔案路徑")
    if file_path is None:
        return

    try:
        hash_str = getOfficeHashString.extract_office_hash(file_path)
        print("John the Ripper 可用哈希:")
        print(hash_str)
    except Exception as e:
        print(f"[错误] {e}")

def office_Crack_PW():
    global CommonPasswordsList
    global BasesWordsList

    file_path = prompt("檔案路徑")
    if file_path is None:
        return

    if user_config["CommonPasswordsList_file"] != "-":
        if os.path.exists(user_config["CommonPasswordsList_file"]):
            with open(user_config["CommonPasswordsList_file"], 'r', encoding='utf-8') as f:
                CommonPasswordsList = [line.strip() for line in f]
        else:
            print("! "+user_config["CommonPasswordsList_file"] + "文件不存在")
    if user_config["BasesWordsList_file"] != "-":
        if os.path.exists(user_config["BasesWordsList_file"]):
            with open(user_config["BasesWordsList_file"], 'r', encoding='utf-8') as f:
                BasesWordsList = [line.strip() for line in f]
        else:
            print("! "+user_config["BasesWordsList_file"] + "文件不存在")

    officeCrackPW.crack_excel_with_generator(file_path,user_config,CommonPasswordsList,BasesWordsList)


def set_Common_Passwords_List():
    global user_config
    file_path = prompt("檔案路徑")
    if file_path is None:
        return
    if os.path.exists(file_path):
        user_config["CommonPasswordsList_file"] = file_path
    else:
        print("! " + file_path + "文件不存在")

def set_Bases_Words():
    global user_config
    file_path = prompt("檔案路徑")
    if file_path is None:
        print("skip")
        return
    if os.path.exists(file_path):
        user_config["BasesWordsList_file"] = file_path
    else:
        print("! " + file_path + "文件不存在")



def show_option():
    global user_config
    # 將 dict 轉成 list of rows
    data = [[key, value if value is not None else "None"] for key, value in user_config.items()]
    headers = ["Options", "value"]

    print_table.print_table(data, headers, align='left')
    print("*不指定 CommonPasswordsList_file, BasesWordsList_file 請填-")

def set_option():
    global user_config

    key = prompt("選項")
    if key is None:
        return
    new_value = prompt("新值")
    if new_value is None:
        return

    if key in user_config.keys():
        user_config[key] = new_value
    else:
        print("!選項不存在")

def load_option():
    global user_config
    try:
        with open(config_filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 如果有參考 dict，檢查 key 是否一致
            if user_config is not None:
                original_keys = set(user_config.keys())
                new_keys = set(data.keys())

                if original_keys != new_keys:
                    print("[✘] JSON 結構不一致")
                    print("原始 keys:", original_keys)
                    print("新載入 keys:", new_keys)
                    return None

            print(f"[✔] 已載入 {config_filename}")
            user_config = data
            return data

    except FileNotFoundError:
        print(f"[✘] 找不到檔案: {config_filename}")
        return None
    except json.JSONDecodeError as e:
        print(f"[✘] JSON 格式錯誤: {e}")
        return None
    except Exception as e:
        print(f"[✘] 載入失敗: {e}")
        return None
def save_option():
    global user_config
    try:
        with open(config_filename, 'w', encoding='utf-8') as f:
            json.dump(user_config, f, ensure_ascii=False, indent=4)
        print(f"[✔] 已儲存至 {config_filename}")
    except Exception as e:
        print(f"[✘] 儲存失敗: {e}")



