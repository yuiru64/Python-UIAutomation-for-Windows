import uiautomation as auto
import uiautomation.event as event

root = auto.GetRootControl()
c = auto.SimpleFind(root, 'Depth:5,RegexName:.*计算器.*')
if not c.Exists(0, 0):
    import subprocess
    subprocess.Popen('calc')
print(c)
def callback(pSender):
    print(pSender)

event.AddFocusChangedEventHandler(callback)
input('Press key to stop')