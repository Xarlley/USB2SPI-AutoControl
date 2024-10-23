'''
    @Author: Liu Shifeng
    @Attention: 
        1. Run this script as an admin of Windows 10/11.
        2. Use pip to install pywinauto Python Lib.
        3. pywin32 Pythyon Lib in pip repository which is required by pywinauto is unusable.
        4. Reinstall pywin32 from release exe insted of pip. Site: https://github.com/mhammond/pywin32/releases
        5. Don't start 'SPI_RW.exe' before this script. This script will automaticly start 'SPI_RW.exe'.
    @What does this script do?
        1. Start and link a process from the source file of 'SPI_RW.exe'.
        2. Access the window of the process.
        3. Automaticly click the button to achieve our goal
'''

from pywinauto import Application
from pywinauto.keyboard import send_keys
import io
import sys
import time

# 通过可执行文件路径启动并连接应用程序
app = Application(backend='win32').start(r'C:\Users\charl\Desktop\projecthjq\SPI_RW_V02\SPI_RW\Debug\SPI_RW.exe')

# 获取主窗口
main_window = app.window()

# 打印所有控件标识符，并截获打印，用于找到每次都会发生变化的控件auto_id
output = io.StringIO()
sys.stdout = output
main_window.print_control_identifiers()
sys.stdout = sys.__stdout__
identifiers_output = output.getvalue()

# 通过查看identifiers_output的内容，来寻找自己需要操作的控件的编号
# print(identifiers_output)

# 先连接设备
connect_button = main_window['连接']
connect_button.click()

def get_edit_content(app, edit_title):
    # 获取主窗口
    main_window = app.window()
    
    # 查找 Edit 控件
    edit_control = main_window.child_window(title=edit_title, control_type="Edit")
    
    # 获取并返回 Edit 中的文本内容
    return edit_control.get_value()

def load_file(app):
    """
    Fail to develop
    Cannot connect to Windows file select function
    So you should read bin file by python code and directly send bin file content to 'Edit' type identifier, then click write-button
    """

def extract_auto_id(output):
    # 将输出按行分割
    lines = output.splitlines()
    auto_id = None

    # 遍历每一行，找到包含 ['Edit4'] 的行(Write Area的输入框控件编号是Edit4)，数字4只是一个demo，实际需要查看identifiers_output找到自己的EditX
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

def read_data_from_bin_file():
    '''
    Define your function that read your bin file.
    '''
    return "aa"

def auto_job(app, identifiers_output):
    """ Demo 1: How to get value of a 'Edit' type identifier
    edit_content = get_edit_content(app, '值：')
    print("Edit 中的内容:", edit_content)
    """

    """ Demo 2: How to click a button
    connect_button = main_window['载入']
    connect_button.click()
    """

    """ Demo 3: How to input content to a 'Edit' type identifier
    auto_id = extract_auto_id(identifiers_output)
    command_edit = main_window.child_window(auto_id=auto_id, control_type="Edit")
    command_edit.set_focus()
    time.sleep(0.5)
    data = read_data_from_bin_file()
    send_keys(data)
    """

    """
    So, you should:
    1. Use demo 1 to get value from '值:', so you can write 'if-else' code to decide what to do
        You can use demo 1 many times, so you can get the value-changed-action to direct what to do
    2. Use demo 2 to click the buttons that you need
        Whatever: '连接，载入，执行，清除，blablabla'
    3. Use demo 3 to get bin file content.
        We cannot connect to Windows file select function, so you should directly write data to Edit
        You should edit the source code of SPI_RW.exe, in order to let Edit-identifier's content can be write by click '执行'
    """

auto_job(app, identifiers_output)