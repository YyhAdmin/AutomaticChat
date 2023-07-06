import win32api, win32gui, win32con
import win32clipboard as clipboard
from pywinauto import Application
import pyautogui
import time

class WindowOperation:
    def __init__(self, className, titleName):
        self.className = className
        self.titleName = titleName

    # 获取窗口句柄，并置于前台
    def get_window(self):
        win = win32gui.FindWindow(self.className, self.titleName)
        win32gui.ShowWindow(win, win32con.SW_MAXIMIZE)
        try:
            win = win32gui.FindWindow(self.className, self.titleName)
        except Exception as e:
            print('请注意：找不到【%s】这个人或群 请激活窗口！' % self.titleName)
        print("找到句柄：%x" % win)
        if win != 0:
            left, top, right, bottom = win32gui.GetWindowRect(win)
            print(left, top, right, bottom)
            win32gui.SetForegroundWindow(win)
            time.sleep(0.5)
        else:
            print('请注意：找不到【%s】这个人或群 请激活窗口!' % self.titleName)
        return win

    # 发送Ctrl+V和Enter键
    def send_m(self, win):
        win32api.keybd_event(17, 0, 0, 0)  # Ctrl键按下
        time.sleep(1)
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, 86, 0)  # V键按下
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # Ctrl键释放
        time.sleep(1)
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # Enter键按下

    # 设置剪贴板内容
    def txt_ctrl_v(self, txt_str):
        clipboard.OpenClipboard()
        clipboard.EmptyClipboard()
        clipboard.SetClipboardData(win32con.CF_UNICODETEXT, txt_str)
        clipboard.CloseClipboard()


class SendMessage(WindowOperation):
    def __init__(self, className, titleName, message):
        super().__init__(className, titleName)
        self.message = message

    # 发送消息
    def send(self):
        """
            使用给定的窗口类名和标题向联系人或组聊天发送消息。
        """

        win = self.get_window()

        # 设置文本到剪贴板
        self.txt_ctrl_v(self.message)

        # 按Ctrl + V将消息粘贴到输入框中
        win32api.keybd_event(17, 0, 0, 0) 
        time.sleep(1)
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, 86, 0) 
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(1)

        # 发送一个回车键来发送消息
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0) 
        win32gui.SendMessage(win, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


class MoveWindow:
    def __init__(self, count):
        self.count = count

    # 拖动窗口
    def drag_window(self):
        start_pos = pyautogui.position()
        for i in range(self.count):
            x = start_pos.x + 2000
            y = start_pos.y

            pyautogui.mouseDown()
            pyautogui.moveTo(x, y, duration=0.3)
            pyautogui.mouseUp()

            pyautogui.moveTo(start_pos.x, start_pos.y, duration=0.3)

class ClickChatButton:
    def __init__(self):
       pass

    def click_chat_button(self):
        # 单击应用程序中具有给定标题的聊天按钮。
        try:
            app = Application().connect(title_re=".*WeChat.*")  # 连接到微信应用
            wechat_main = app.window(title_re=".*WeChat.*")  # 获取微信主窗口

   
            chat_button = wechat_main.child_window(control_id=12345)
            chat_button.click()  # 点击聊天按钮

            # 将微信窗口移动到屏幕左上角，并设置大小为800x600
            wechat_main.move_window(x=0, y=0, width=800, height=600)
        except Exception as e:
            print(f"Failed to click the chat button: {e}")


if __name__ == '__main__':
    # 输入拖拽次数
    count = int(input('请输入拖拽次数:'))
    # 创建MoveWindow对象，并拖拽窗口
    mv = MoveWindow(count)
    mv.drag_window()

    # 设置要发送消息的联系人或群聊名字和消息内容
    titles = ['杨雨航']
    message = '测试消息'

    # 对每一个联系人或群聊发送消息
    for title in titles:
        sm = SendMessage('ChatWnd', title, message)
        sm.send()

    # 创建ClickChatButton对象，并点击聊天按钮
    # cb = ClickChatButton()
    # cb.click_chat_button()