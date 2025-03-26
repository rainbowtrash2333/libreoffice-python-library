# testing
import uno
import subprocess
import socket
import sys

from libreoffice_py import document
from calendar import Calendar
from datetime import date
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append('.')
today = date.today()
NoConnectException = uno.getClass('com.sun.star.connection.NoConnectException')

def __test_save_ods():
    calc = document.Calc()
    sheet = calc.get_sheets()[0]
    range = sheet.get_range(0, 0, 1, 1)
    data = [
        ['Hello', 'World'],
        [0, 1]
    ]
    range.set_data(data)
    calc.save('./saves/testfile.ods')
    calc.desktop.terminate()

def __test_save_pdf():
    today = date.today()

    calc = document.Calc()
    sheet = calc.get_sheets()[0]

    month = 4
    cal = Calendar(firstweekday=0)
    monthdates = [
        f'{d[0]}.{d[1]}.{d[2]}' for d in cal.itermonthdays4(today.year, month)
        if d[3] not in (5, 6) and d[1] == month
    ]

    data = [
        ['Firm', 'BRIGHT', '', '', ''],
        ['Address', 'Street no.', '', '', ''],
        ['Employee', 'Jack Sprot', '', '', ''],
        ['Date', 'Arrive at', 'Sign at arrival', 'Leave at', 'Sign at arrival']
    ]

    for d in monthdates:
        data.append([d, '', '08:00', '', '18:00'])

    for line in data:
        print(line)
    range = sheet.get_range(0, 0, len(data[0]) - 1, len(data) - 1)

    range.set_data(data)
    calc.save('./saves/testfile.pdf', 'pdf')
    calc.desktop.terminate()

def __test_start_libreoffice():
    start_command1 = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "--headless",
        "--invisible",
        "--nocrashreport",
        "--nodefault",
        "--nofirststartwizard",
        "--nologo",
        "--norestore",
        "--accept=socket,host=localhost,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext"
    ]
    lo_proc = subprocess.Popen(start_command1, shell=False)
    return lo_proc


def __shutdown_soffice():
    try:
        # 初始化 UNO 连接（需与启动时的 --accept 参数匹配）
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)

        # 连接到运行中的 soffice 实例
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

        # 获取桌面服务并调用终止
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        desktop.terminate()  # 优雅关闭
        print("LibreOffice headless 服务已关闭。")

    except NoConnectException:
        print("连接失败：确保 soffice 正在运行且端口 2002 可用。")
    except Exception as e:
        print(f"错误：{e}")


def __test_is_start(port=2002):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(1)  # 设置超时时间
        result = s.connect_ex(('127.0.0.1', port))
        s.close()
        return result == 0  # 返回 True 表示端口已监听
    except Exception as e:
        return False


if __name__ == '__main__':
    # __shutdown_soffice()
    # __test_save_ods()
    __test_save_pdf()
    # lo_proc = __test_start_libreoffice()
    # print(__test_is_start())
