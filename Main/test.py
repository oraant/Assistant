import threading
import win32com.client as wincl
import pyttsx3

engine = pyttsx3.init()

speak = wincl.Dispatch("SAPI.SpVoice")
t = threading.Event()


def s(data=""):
    global t
    t.set()
    if data == "": data = """This is a story of two tribal Armenian boys who belonged to the 
         Garoghlanian tribe. """
    s = speak.Speak(data)


t1 = threading.Thread(target=s)

t2 = threading.Thread(target=s, args=("o o o o o o o o o",))
t2.start()
t1.start()

exit()
# ------------------------------------------------------------------------------------------------------------------

import pyttsx3;
engine = pyttsx3.init();
engine.say("I will speak this text");
engine.say("I will speak this text");
engine.runAndWait() ;
engine.stop()

exit()
# ------------------------------------------------------------------------------------------------------------------


import winsound
import win32com.client
import time
import yaml
import random
import threading

speak = win32com.client.Dispatch('SAPI.SPVOICE').Speak

# 公共变量
l = [1,2,3,4,5,6,7,8,9,9]
i = 22

def f(a):
    print(a)

# ------------------------------------------------------------------------------------------------------------------


l.append([10]*i)
print(l)

for i in range(5):
    print(i)



t = threading.Thread(target=f, args=("haha"))
t.start()
t.join()

exit()
with open('./nodes.yaml', encoding="utf-8") as f:
    y = yaml.load(f)
    y[0]["count"] = 999
    print(y)
    yaml.dump(y, f, Dumper=yaml.RoundTripDumper)
