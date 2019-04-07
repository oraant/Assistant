import threading
import win32con
import win32com.client
import ctypes
import ctypes.wintypes
import random
import time
import os
from ruamel.yaml import YAML


# 方法缩写
speak = win32com.client.Dispatch('SAPI.SPVOICE').Speak  # Windows朗读设备，只能阻塞朗读
user32 = ctypes.windll.user32  # Windows下的user32.dll
curPath = os.path.dirname(os.path.realpath(__file__))
yamlPath = os.path.join(curPath, "nodes.yaml")

# 全局变量
config = []  # 配置表中的配置
seconds = 0  # 还有多少秒退出
node = ""  # 最终选择的内容

# 注册热键
class Hotkey(threading.Thread):  # 创建一个Thread.threading的扩展类

    global seconds

    HotKeys = {
        # ID，组合键，键盘键
        # Ctrl + 0，加10分钟； Ctrl + 5，加5分钟
        10 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD0, "+", 10*60),
        11 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD1, "+", 1*60),
        12 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD2, "+", 2*60),
        13 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD3, "+", 3*60),
        14 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD4, "+", 4*60),
        15 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD5, "+", 5*60),
        16 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD6, "+", 6*60),
        17 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD7, "+", 7*60),
        18 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD8, "+", 8*60),
        19 : (win32con.MOD_CONTROL, win32con.VK_NUMPAD9, "+", 9*60),

        # Alt + 0，减10分钟； Alt + 5，减5分钟
        20 : (win32con.MOD_ALT, win32con.VK_NUMPAD0, "-", 10*60),
        21 : (win32con.MOD_ALT, win32con.VK_NUMPAD1, "-", 1*60),
        22 : (win32con.MOD_ALT, win32con.VK_NUMPAD2, "-", 2*60),
        23 : (win32con.MOD_ALT, win32con.VK_NUMPAD3, "-", 3*60),
        24 : (win32con.MOD_ALT, win32con.VK_NUMPAD4, "-", 4*60),
        25 : (win32con.MOD_ALT, win32con.VK_NUMPAD5, "-", 5*60),
        26 : (win32con.MOD_ALT, win32con.VK_NUMPAD6, "-", 6*60),
        27 : (win32con.MOD_ALT, win32con.VK_NUMPAD7, "-", 7*60),
        28 : (win32con.MOD_ALT, win32con.VK_NUMPAD8, "-", 8*60),
        29 : (win32con.MOD_ALT, win32con.VK_NUMPAD9, "-", 9*60),
    }

    # 当相关热键被按下时，执行相关的命令，并进行相关提示
    def handler(self, id):
        global seconds

        operation = self.HotKeys.get(id)[2]
        number = self.HotKeys.get(id)[3]
        if operation == "+":
            seconds = seconds + number
        else:
            seconds = seconds - number
            if seconds < 0: seconds = 9

        statement = "您更新了倒计时时间，现在时间为：" + seconds2str(seconds)
        print(statement)
        speak(statement)

    # 批量注册热键
    def register(self):
        for id, (modifier, key, a, b) in self.HotKeys.items():
            if not user32.RegisterHotKey(None, id, modifier, key):
                print("热键： %s + %s 注册失败！" % (modifier, key))

    # 批量取消注册过的热键，必须得释放热键，否则下次就会注册失败
    def unregister(self):
        for id in self.HotKeys.keys():
            user32.UnregisterHotKey(None, id)

    # 线程执行
    def run(self):

        self.register()

        # 循环检测热键是否被按下
        try:
            msg = ctypes.wintypes.MSG()
            while user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:  # 当监测到有热键被按下时
                    self.handler(msg.wParam)  # msg.wParam 就是热键注册的ID
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageA(ctypes.byref(msg))
        finally:
            self.unregister()


# ------ 存取配置 -------------------------------------------------------------------------

# 从配置表中获取配置，将配置文件中的配置列表，读取到config变量中
def get_config():
    global config, yamlPath
    yaml = YAML(typ='safe')
    with open(yamlPath, encoding="utf-8") as f:
        config = yaml.load(f)

# 向配置表中写入配置：根据当前的Node，将Config中的count改成相应的值，并写入配置文件
def set_config():
    global config, node, yamlPath

    try: # 尝试获取下标，获取成功说明是正常选择，否则是自定义内容
        subscript = get_subscript(config, node)
        config[subscript]["count"] += 1
    except:
        print("项目%s为自定义内容，无法增加权重" % (node))
        return

    # 根据选择的节点，更改其权重
    with open(yamlPath, 'w', encoding="utf-8") as f:
        yaml = YAML()
        yaml.default_flow_style = False
        yaml.dump(config, f)

    # 验证更改后的权重是否生效
    get_config()
    print("项目%s的权重已增加为：%d" % (node, config[subscript]["count"]))

# 根据名字查找其在配置表中的下标
def get_subscript(nodes, node):
    for i, value in enumerate(nodes):
        if value["name"] == node: return i
    raise ValueError


# ------ 获取节点 -------------------------------------------------------------------------

# 从配置列表中，按照一定的方式抽取几张卡片
def get_nodes(samples = 3, weight = True, switch=True):  # 取样数量，是否按照权重取样，是否按照开关取样

    global config
    nodes = []

    for node in config:
        name, count, active = node["name"], node["count"], node["active"]
        if switch == True and active == False: continue
        count = count+1 if weight == True else 1
        for i in range(count): nodes.append(name)

    return random.sample(nodes, samples)

# 展示列表
def show_nodes(nodes):
    print("\n当前抽取的卡片如下，请用数字选择卡片，用0重新抽取卡片，其他选项请打字输入：")
    for i, v in enumerate(nodes):
        print("%d - %s" % (i+1, str(v)))

# 用户选择节点，或者自己编写新节点，最后将值存放到Node中
def choice_node():
    global node

    while 1:
        get_config()
        nodes = get_nodes()
        show_nodes(nodes)
        response = input()

        if response.isdigit():
            choice = int(response)
            if choice == 0:
                continue
            else:
                node = nodes[choice-1]
                break
        else:
            node = response
            break

    set_config()
    print("您选择了：" + node)

# ------ 获取时间 -------------------------------------------------------------------------

# 将20m转换为1200秒
def str2seconds(string):
    print(string)
    if string.isdigit(): string += "m"
    num = int(string[:-1])
    suffix = string[-1]

    if suffix == "h":
        return num * 3600
    elif suffix == "m" or "":
        return num * 60
    elif suffix == "s":
        return num
    else:
        raise ValueError

# 将1200秒转换为20分钟
def seconds2str(number):
    if number > 3600:
        return "%.1f" % (number/3600) + "小时"
    elif number > 60:
        return "%.1f" % (number/60) + "分钟"
    elif number >= 0:
        return "%.1f" % (number) + "秒"
    else:
        raise ValueError

# 让客户输入倒计时多久，默认三十分钟
def choice_time():
    global seconds

    while 1:
        response = input("\n请设置多少分钟后使用此卡片(默认为30分钟)：\n")
        if response == "" : seconds = 1800; return
        try:
            seconds = str2seconds(response)
            break
        except:
            continue

    print("您将时间设置为%s后\n" % seconds2str(seconds))

# ------ 计时提醒 -------------------------------------------------------------------------

# 设置不同的有趣的提醒词，并且给提醒者赋予
def remind(node, time):
    statements = [  # todo: 写到配置文件里。同步的github上。拆分成多个文件，包括弄配置的类、弄语音的类(获取token之类的，若离线则使用微软自带)、注册热键的类等。Ctrl+.报时功能
        # 名称在前，时间在后
        "衍主在上：距离%s还有%s哦" % (node, time),
        "%s这种事，衍主真的喜欢吗，再等%s吧" % (node, time),
        "想%s吗？再等%s就可以啦" % (node, time),
        "想%s还得再等%s哦，静姝跟您一起等哦" % (node, time),
        "%s喽，可惜还得再等%s哦，小姝也想试试呢" % (node, time),

        # 时间在前，名称在后
        "衍主衍主，偷偷告诉你，还有%s，就可以%s啦" % (time, node),
        "当当当，还有%s，就可以%s啦，开心吗" % (time, node),
        "加油哦，还有%s，就可以%s啦，开心吗" % (time, node),
        "笨笨的静姝提醒您，还有%s，就可以%s喽" % (time, node),
        "随静姝又来啦，还有%s，就可以%s喽" % (time, node),
    ]

    statement = random.sample(statements, 1)
    speak(statement)

# 在不同的时间段，提醒不同的内容
def checkpoint():
    global seconds
    if seconds == 7200:
        remind(node, "两个小时")
    if seconds == 3600:
        remind(node, "一个小时")
    if seconds == 1800:
        remind(node, "30分钟")
    if seconds == 1200:
        remind(node, "20分钟")
    if seconds == 600:
        remind(node, "10分钟")
    if seconds == 300:
        remind(node, "5分钟")
    if seconds == 180:
        remind(node, "3分钟")
    if seconds == 120:
        remind(node, "2分钟")
    if seconds == 60:
        remind(node, "1分钟")
    if seconds == 30:
        remind(node, "30秒")
    if seconds == 20:
        remind(node, "20秒")
    if seconds == 10:
        speak("只剩十秒啦，快要倒计时喽")
    if seconds == 5:
        speak("5！")
    if seconds == 4:
        speak("4！")
    if seconds == 3:
        speak("3！")
    if seconds == 2:
        speak("2！")
    if seconds == 1:
        speak("1！")
    if seconds == 0:
        speak("当当当当！时间到啦！快去%s吧！" % node)

# 倒计时，一秒减一次
def countdown():
    global seconds
    while seconds >= 0:
        checkpoint()
        time.sleep(1)
        seconds = seconds - 1

# ------ 正式执行 -------------------------------------------------------------------------

# 注册热键
hotkey = Hotkey()
hotkey.setDaemon(True)
hotkey.start()

# 执行程序
choice_node()
choice_time()
countdown()
