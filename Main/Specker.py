import winsound
import random
import win32com.client
import time

speak = win32com.client.Dispatch('SAPI.SPVOICE').Speak
sleep = lambda x: time.sleep(x*60)

nodes = [
    # 娱乐
    "谜语", "灯谜", "急转弯",
    "OW技能", "火影技能", #
    "画符号", "掐手决", "摆姿势", "念咒语",
    "组词", "扔石子游戏", "叠纸鹤", "手指勾线", "积木", "涂鸦", "涂色", "手工",
    "百度网盘归档分类",
    "相声",
    "下载书籍", "下载Unity插件",
    "全平台打卡软件",
    "博馆",

    # 水果零食
    "冰糕", "酸奶", "牛奶", "巧克力", "火腿", "辣条",
    "喝咖啡", "喝两口水", "喝瓶可乐",
    "草莓", "香蕉", "火龙果", "橘子", "西瓜", "樱桃", "猕猴桃", "桃子", "葡萄", "菠萝", "甘蔗", 

    # 通用学习
    "日语", "英语", "法语", "俄语", "世界语", # 语言
    "跑步", "散步", "俯卧撑", "伸展吐纳", # 锻炼
    "经典名著", "诗经", "论语", "诗词歌赋", "散文", "成语典故", # 文学
    "自然科学", "财富经济", "政治军事", "时尚设计", # 阅读
    "早睡", "早起", "午睡", "打滚", "蒙头", "大躺", "小憩", "大脑放空", # 休息

    # 专业学习
    "高数", "线代", "概率", # 高数
    "系统", "计组", "算法", "网络", # 专业课
    "单词", "阅读", "长难句", "写作", # 英语
    "马哲", "毛概", "思修", # 政治
    "刷题", "书籍", # 技术

    # 通用创作
    "散文", "写诗", "写小说", "写日记", "记录心情", "创作" # 文档
    "手绘", "板绘", "漫画", "挑战", "赏析", "名画", "动画", "书法", # 图片
    "钢琴", "葫芦丝", "古筝", "名曲", "简谱", "相声", # 音频
    "采访", "动画", # 视频
    "建模", "动画", "渲染", # 三维
]




# 抽三张卡片，用户从三张里自选一张
choice = 0
while choice == 0:
    node = random.sample(nodes, 3)
    choice = input("当前抽取的卡片为：" + str(node) + "，\n请用1、2、3选择卡片，用0重新抽取卡片，其他选项请打字输入：\n")
    if choice in "123":
        choice = int(choice)
        node = node[choice - 1]
    elif choice == "0":
        choice = 0
    else:
        node = choice


# 让客户输入倒计时多久，默认三十分钟
minutes = input("您当前的卡片为：" + node + "。\n请选择多少分钟后使用此卡片(默认为30分钟)：\n")
minutes = int(minutes) if minutes != "" else 30


# 倒计时还有多少分钟的时候，报一下时间和选择的项目
def waitm_and_speak(waitm, word):
    global minutes
    if minutes < waitm: return
    time.sleep((minutes - waitm)*60)
    speak("再过%d分钟，您就可以%s啦：" % (waitm, word))
    minutes = waitm

waitm_and_speak(120, node)
waitm_and_speak(60, node)
waitm_and_speak(30, node)
waitm_and_speak(20, node)
waitm_and_speak(10, node)
waitm_and_speak(5, node)
waitm_and_speak(3, node)
waitm_and_speak(2, node)


# 开始以秒为单位进行计算
seconds = minutes*60

# 倒计时还有多少秒钟的时候，报一下时间和选择的项目
def waits_and_speak(waits, word):
    global seconds
    if seconds < waits: return
    time.sleep(seconds - waits)
    speak("再过%d秒，%s：" % (waits, word))
    seconds = waits

waits_and_speak(60, node)
waits_and_speak(30, node)
waits_and_speak(20, node)
waits_and_speak(10, node)

time.sleep(5)
for i in [5, 4, 3, 2, 1]:
    speak(str(i))
    time.sleep(1)

winsound.PlaySound(node, winsound.SND_ASYNC)
time.sleep(2) # 为了让声音响完