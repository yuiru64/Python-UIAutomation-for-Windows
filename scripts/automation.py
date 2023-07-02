#!D:\Environment\Python3\python3.9.exe
# -*- coding:utf-8 -*-
import sys
import time

import uiautomation as auto


def usage():
    auto.Logger.ColorfullyWrite("""usage
<Color=Cyan>-h, --help</Color>      show command <Color=Cyan>help</Color>
<Color=Cyan>-s, --sleep</Color>      delay <Color=Cyan>time</Color>, default 3 seconds, begin to enumerate after Value seconds, this must be an integer
        you can delay a few seconds and make a window active so automation can enumerate the active window
<Color=Cyan>-d, --depth</Color>      enumerate tree <Color=Cyan>depth</Color>, this must be an integer, if it is null, enumerate the whole tree
<Color=Cyan>-r, --root</Color>      enumerate from <Color=Cyan>root</Color>:Desktop window, if it is null, enumerate from foreground window
<Color=Cyan>-f, --foucs</Color>      enumerate from <Color=Cyan>focused</Color> control, if it is null, enumerate from foreground window
<Color=Cyan>-c, --cursor</Color>      enumerate the control under <Color=Cyan>cursor</Color>, if depth is < 0, enumerate from its ancestor up to depth
<Color=Cyan>-a, --ancestor</Color>      show <Color=Cyan>ancestors</Color> of the control under cursor
<Color=Cyan>-n, --showAllName</Color>      show control full <Color=Cyan>name</Color>, if it is null, show first 30 characters of control's name in console,
        always show full name in log file @AutomationLog.txt
<Color=Cyan>-p</Color>      show <Color=Cyan>process id</Color> of controls
<Color=Cyan>-l, --long</Color>      show more detailed information
<Color=Cyan>-x, --xfind</Color>      find control by syntax 'Name:xxx/RegexName:xxx'
<Color=Cyan>-t, --timeout</Color>      timeout to search control

if <Color=Red>UnicodeError</Color> or <Color=Red>LookupError</Color> occurred when printing,
try to change the active code page of console window by using <Color=Cyan>chcp</Color> or see the log file <Color=Cyan>@AutomationLog.txt</Color>
chcp, get current active code page
chcp 936, set active code page to gbk
chcp 65001, set active code page to utf-8

examples:
automation.py -t3
automation.py -t3 -r -d1 -n
automation.py -c -t3

xfind syntax:
expression: <term>/<term>/...
<term>: <attr_term> | <func_term>
<attr_term>: <attr>:value,<attr>:value,...
<func_term>: @child:<int> | @first | @last | @next | @prev

Supported Attr:
searchDepth: int, max search depth from searchFromControl.
foundIndex: int, starts with 1, >= 1.
searchInterval: float, wait searchInterval after every search in self.Refind and self.Exists, the global timeout is TIME_OUT_SECOND.
searchProperties: defines how to search, the following keys can be used:
    ControlType: int, a value in class `ControlType`.
    ClassName: str.
    AutomationId: str.
    Name: str.
    SubName: str, a part str in Name.
    RegexName: str, supports regex using re.match.
        You can only use one of Name, SubName, RegexName in searchProperties.
    Depth: int, only search controls in relative depth from searchFromControl, ignore controls in depth(0~Depth-1),
        if set, searchDepth will be set to Depth too.
    Compare: Callable[[Control, int], bool], custom compare function(control: Control, depth: int) -> bool.

""", writeToFile=False)


def SimpleEnumAndLogControl(control, depth, startIndent, showFullName: bool=False):
    for c, d in auto.WalkControl(control, True, depth):
        auto.Logger.Write("    " * (startIndent + d))
        auto.Logger.Write("ControlTypeName: ")
        auto.Logger.Write(c.ControlTypeName, auto.ConsoleColor.Green)
        auto.Logger.Write(", Name: ")
        auto.Logger.Write(c.Name[:min(len(c.Name), 30)] if not showFullName else c.Name, auto.ConsoleColor.Green)
        auto.Logger.Write(", ClassName: ")
        auto.Logger.Write(c.ClassName, auto.ConsoleColor.Green)
        auto.Logger.Write("\n")

def SimpleEnumAndLogControlAncestors(control, showFullName: bool=False):
    lists = []
    while control:
        lists.insert(0, control)
        control = control.GetParentControl()
    for i, c in enumerate(lists):
        auto.Logger.Write("    " * i)
        auto.Logger.Write("ControlTypeName: ")
        auto.Logger.Write(c.ControlTypeName, auto.ConsoleColor.Green)
        auto.Logger.Write(", Name: ")
        auto.Logger.Write(c.Name[:min(len(c.Name), 30)] if not showFullName else c.Name, auto.ConsoleColor.Green)
        auto.Logger.Write(", ClassName: ")
        auto.Logger.Write(c.ClassName, auto.ConsoleColor.Green)
        auto.Logger.Write("\n")

def main():
    import getopt
    auto.Logger.Write('UIAutomation {} (Python {}.{}.{}, {} bit)\n'.format(auto.VERSION, sys.version_info.major, sys.version_info.minor, sys.version_info.micro, 64 if sys.maxsize > 0xFFFFFFFF else 32))
    options, args = getopt.getopt(sys.argv[1:], 'hrfcanpd:s:x:lt:',
                                  ['help', 'root', 'focus', 'cursor', 'ancestor', 'showAllName', 'depth=',
                                   'sleep=', "xfind", "long", "timeout="])
    root = False
    focus = False
    cursor = False
    ancestor = False
    foreground = True
    showAllName = False
    depth = 0xFFFFFFFF
    seconds = 3
    showPid = False
    long = False
    findMatcher = None
    for (o, v) in options:
        if o in ('-h', '--help'):
            usage()
            exit(0)
        elif o in ('-r', '--root'):
            root = True
            foreground = False
        elif o in ('-f', '--focus'):
            focus = True
            foreground = False
        elif o in ('-c', '--cursor'):
            cursor = True
            foreground = False
        elif o in ('-a', '--ancestor'):
            ancestor = True
            foreground = False
        elif o in ('-n', '--showAllName'):
            showAllName = True
        elif o in ('-p', ):
            showPid = True
        elif o in ('-d', '--depth'):
            depth = int(v)
        elif o in ('-s', '--sleep'):
            seconds = float(v)
        elif o in ('-l', '--long'):
            long = True
        elif o in ('-t', '--timeout'):
            auto.setGlobal("TIME_OUT_SECOND", int(v))
        elif o in ('-x', '--xfind'):
            findMatcher = v

    if seconds > 0:
        auto.Logger.Write('please wait for {0} seconds\n\n'.format(seconds), writeToFile=False)
        time.sleep(seconds)
    auto.Logger.Log('Starts, Current Cursor Position: {}'.format(auto.GetCursorPos()))
    control = None

    if root:
        control = auto.GetRootControl()
    if focus:
        control = auto.GetFocusedControl()
    if cursor:
        control = auto.ControlFromCursor()
        if depth < 0:
            while depth < 0 and control:
                control = control.GetParentControl()
                depth += 1
            depth = 0xFFFFFFFF
    if ancestor:
        if findMatcher:
            control = auto.SimpleFind(control, findMatcher)
        control = auto.ControlFromCursor()
        if control:
            if long:
                auto.EnumAndLogControlAncestors(control, showAllName, showPid)
            else:
                SimpleEnumAndLogControlAncestors(control, showAllName)
        else:
            auto.Logger.Write('IUIAutomation returns null element under cursor\n', auto.ConsoleColor.Yellow)
    else:
        indent = 0
        if not control:
            control = auto.GetFocusedControl()
            controlList = []
            while control:
                controlList.insert(0, control)
                control = control.GetParentControl()
            if len(controlList) == 1:
                control = controlList[0]
            else:
                control = controlList[1]
                if foreground:
                    indent = 1
                    auto.LogControl(controlList[0], 0, showAllName, showPid)
        if findMatcher:
            control = auto.SimpleFind(control, findMatcher)
        if long:
            auto.EnumAndLogControl(control, depth, showAllName, showPid, startDepth=indent)
        else:
            SimpleEnumAndLogControl(control, depth, indent, showAllName)
    auto.Logger.Log('Ends\n')


if __name__ == '__main__':
    auto.Logger.DeleteLog()
    auto.Debug()
    main()
