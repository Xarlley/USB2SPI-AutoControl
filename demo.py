'''
    @Author: Liu Shifeng
    @Attention: 
        1. Run this script as an admin of Windows 10/11.
        2. Use pip to install pywinauto Python Lib.
        3. pywin32 Pythyon Lib in pip repository which is required by pywinauto is unusable.
        4. Reinstall pywin32 from release exe insted of pip. Site: https://github.com/mhammond/pywin32/releases
        5. Don't start 'USB2SPI.exe' before this script. This script will automaticly start 'USB2SPI.exe'.
    @What does this script do?
        1. Start and link a process from the source file of 'USB2SPI.exe'.
        2. Access the window of the process.
        3. Automaticly click the button to achieve our goal:
            A. 'Connect' - Connect to hardware.
            B. Loop of Write-text-to-Write_Area and 'Do write' - Our main job, loading data from bin file and write data to device.
'''

from pywinauto import Application
from pywinauto.keyboard import send_keys
import io
import sys
import re
import time
import os

# 通过可执行文件路径启动并连接应用程序
app = Application(backend='uia').start(r'C:\Program Files (x86)\ONYIB\USB2SPI\bin\USB2SPI.exe')

# 获取主窗口
main_window = app.window()

# 打印所有控件标识符，并截获打印，用于找到每次都会发生变化的控件auto_id
output = io.StringIO()
sys.stdout = output
main_window.print_control_identifiers()
sys.stdout = sys.__stdout__
identifiers_output = output.getvalue()

def extract_auto_id(output):
    # 将输出按行分割
    lines = output.splitlines()
    auto_id = None

    # 遍历每一行，找到包含 ['Edit4'] 的行(Write Area的输入框控件编号是Edit4)
    for i, line in enumerate(lines):
        if "['Edit4']" in line:
            # 找到 ['Edit4'] 后，检查下一行是否包含 auto_id
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                match = re.search(r'auto_id="(\d+)"', next_line)  # 正则匹配 auto_id
                if match:
                    auto_id = match.group(1)
                    print(f"Found auto_id: {auto_id}")
                    return auto_id
            break
    
    if not auto_id:
        print("auto_id not found for Edit4")
    return auto_id

auto_id = extract_auto_id(identifiers_output)

if auto_id == None:
    print('auto_id not found')
    os.exit(0)

# 先连接设备
main_window['Connect'].click()
print("11")
edit_control = main_window.child_window(control_type="Edit", found_index=0)

def read_data_from_bin_file()
    '''
    Define your function that read your bin file.
    '''
    return "aa"

# 每隔 20 秒点击按钮的函数
def auto_click_buttons():
    while True:

        # 使用 send_keys() 方法将文件路径输入到文件名文本框中
        time.sleep(2)
        command_edit = main_window.child_window(auto_id=auto_id, control_type="Edit")
        command_edit.set_focus()
        time.sleep(0.5)
        data = read_data_from_bin_file()
        send_keys("aa")

        # 点击 'Write to device' 按钮
        main_window['Do write'].click()
        print('Clicked Write to device button')
        
        # 每 20 秒重复一次
        time.sleep(20)

# 调用函数开始自动点击
auto_click_buttons()