# uiautomation

[[中文](readme_cn.md)] [[Englinsh](readme.md)]

uiautomation is a module developed by yinkaisheng for personal use in his spare time. As of 2023, the original author has not updated it for almost two years. This is a fork version of uiautomation that adds some additional features to the original module.

uiautomation encapsulates the Microsoft [UIAutomation API](https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomation) and supports automation for Win32, MFC, WPF, Modern UI (Metro UI), Qt, IE, Firefox (version <= 56 or >= 60, Firefox 57 is the first Rust development version, and personal tests have shown that the first few Rust development versions do not support it), Chrome, and applications based on Electron (Chrome browser and Electron applications require the "--force-renderer-accessibility" launch parameter to support UIAutomation).

In addition, uiautomation also provides screenshot functionality, which is sourced from [UIAutomationClient](https://github.com/yinkaisheng/UIAutomationClient).

## Features

- [x] Command-line control search
- [x] Multiple conditions for control matching and traversal
- [x] Screen capture
- [x] Hotkey binding
- [x] Event listening
- [x] Support for running in different threads
- [ ] ...

## Installation Requirements

### Version Compatibility

The latest version of uiautomation2.0 only supports Python 3, with dependencies on the "comtypes" and "typing" packages. However, versions 3.7.6 and 3.8.1 should not be used, as comtypes does not work properly in these versions ([issue](https://github.com/enthought/comtypes/issues/202)). Use the following command to install the dependencies:

```
pip install comtypes
```

If you are using code from versions prior to uiautomation2.0, please refer to the [API changes](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows/blob/master/API%20changes.txt) to modify your code. uiautomation supports running on Windows desktop systems starting from Windows XP SP3 or later.

### Requirements for Older Windows Systems

If you are using Windows XP, make sure that the system directory contains the file "UIAutomationCore.dll". If it is not present, you need to install the patch **[KB971513](https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation)** to support UIAutomation.

### Administrator Privileges Requirement

When using uiautomation on Windows 7 or later versions of Windows, Python needs to be run with administrator privileges; otherwise, many functions of uiautomation may fail or throw exceptions. Alternatively, you can run cmd.exe with administrator privileges and then call Python in the cmd window. The title of the cmd window should indicate **Administrator**.

After installing "uiautomation" using "pip install uiautomation", a file called "automation.py" will be available in the Scripts directory of Python (e.g., C:\Python37\Scripts), or you can use the "automation.py" file from the root directory of the source code. "automation.py" is a script used to enumerate the control tree structure.

## Get started Guide

### Searching with Command Line

Run '**automation.py -h**' to view the command help. When writing automation code, refer to the output of this command to write corresponding code.
![help](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\uiautomation-h-16858087959281.png)

Understand the meanings of the parameters in the above image and run the following commands to see the program's execution results.
**automation.py -t 0**,   Print all controls of the currently active window.
**automation.py -r -d 1 -t 0**, Print the desktop (root control of the tree) and its first-level child windows (TopLevel windows).
**automation.py -xfind Depth:1,RegexName:.\*Calculator.\*/@first/@last/@parent/@next/@prev/@child:3**, Search for controls that match the regular expression `.*Calculator.*` in nodes with a depth of 1. Take the first child node, then take the last child node of the new node, then take the parent node of the new node, then take the next sibling of the new node, then take the previous sibling of the new node, and finally take the third child node of the new node.

![top level windows](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\automation_toplevels.png)

automation.py displays some properties of each control in the control tree and the patterns supported by the control.

According to Microsoft UIAutomation API, a specific type of control must support or choose to support a certain pattern, as shown in the image below:

![patterns](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\control_pattern.png)

Refer to [Control Pattern Mapping for UI Automation Clients](https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpatternmapping) to view the complete table of control-pattern support.

uiautomation encapsulates various controls and patterns from Windows UIAutomation.

Control types include ButtonControl, TextControl, TreeControl, and more.

Pattern types include ExpandCollapsePattern, InvokePattern, and more.

In actual use, Control or Pattern objects are used to retrieve control information or operate on controls.

uiautomation searches for controls in the control tree from top to bottom based on the control properties you provide.

Assuming the control tree is as follows:

```
root(Name='Desktop', Depth=0)
    window1(Depth=1)
        control1-001(Depth=2)
        control1-...(Depth=2)
        ...
        control1-100(Depth=2)
    window2(Name='window2', Depth=1)
        control2-1(Depth=2)
            control2-1-001(Depth=3)
            control2-1-...(Depth=3)
            ...
            control2-1-100(Depth=3)
        control2-2(Depth=2)
        control2-3(Depth=2)
        control2-4(Name='2-4', Depth=2)
            editcontrol(Name='myedit1', Depth=3)
            editcontrol(Name='myedit2', Depth=3
```

### Searching with Code

If you want to find an EditControl named "myedit2" and type in it, you can write the following code:

```python
uiautomation.EditControl(searchDepth=3, Name='myedit2').SendKeys('hi')
```

However, this code is not very efficient because there are many controls in the control tree, and the EditControl you are looking for is at the end of the tree. It takes over 200 iterations to find this EditControl by searching the entire control tree from the root.

To speed up the search, you can use layered searching and specify the search depth. This way, you can find the control in just a few iterations. Here's an example:

```python
window2 = uiautomation.WindowControl(searchDepth=1, Name='window2')  # search 2 times
sub = window2.Control(searchDepth=1, Name='2-4')  # search 4 times
edit = sub.EditControl(searchDepth=1, Name='myedit2')  # search 2 times
edit.SendKeys('hi')
```

In this code, we first search for the window named "window2" in the first layer of the desktop. It takes 2 iterations to find it. Then, we search for the control named "2-4" in the first layer of "window2". It takes 4 iterations to find it. Finally, we search for the EditControl named "myedit2" in the first layer of "2-4". It takes 2 iterations to find it. In total, it takes 8 iterations to find the control.

You can also combine the above four lines of code into one line:

```python
uiautomation.WindowControl(searchDepth=1, Name='window2').Control(searchDepth=1, Name='2-4').EditControl(searchDepth=1, Name='myedit2').SendKeys('hi')
```

Let's take a look at an example with the Notepad program. Run "notepad.exe" and then run "automation.py -t 3" to switch to Notepad and make it the currently active window. After 3 seconds, "automation.py" will print all the controls in Notepad and save them to the log file "@AutomationLog.txt".

On my computer, the output looks like this:

```
ControlType: PaneControl    ClassName: #32769    Name: Desktop    Depth: 0    **(Desktop window, root control)**
    ControlType: WindowControl    ClassName: Notepad    Depth: 1    **(Top-level window, Notepad window)**
        ControlType: EditControl    ClassName: Edit    Depth: 2
            ControlType: ScrollBarControl    ClassName:     Depth: 3
                ControlType: ButtonControl    ClassName:     Depth: 4
                ControlType: ButtonControl    ClassName:     Depth: 4
            ControlType: ThumbControl    ClassName:     Depth: 3
        ControlType: TitleBarControl    ClassName:     Depth: 2
            ControlType: MenuBarControl    ClassName:     Depth: 3
                ControlType: MenuItemControl    ClassName:     Depth: 4
            ControlType: ButtonControl    ClassName:     Name: Minimize    Depth: 3
            ControlType: ButtonControl    ClassName:     Name: Maximize    Depth: 3
            ControlType: ButtonControl    ClassName:     Name: Close    Depth: 3
```

Run the following code:

```python
# -*- coding: utf-8 -*-
import subprocess
import uiautomation as auto

print(auto.GetRootControl())
subprocess.Popen('

notepad.exe')
# First, find the window control of the Notepad program from the first layer of the desktop's child controls.
notepadWindow = auto.WindowControl(searchDepth=1, ClassName='Notepad')
print(notepadWindow.Name)
notepadWindow.SetTopmost(True)
# Find the first EditControl among all descendants of the notepadWindow control. Since the EditControl is the first child control, you don't need to specify the depth.
edit = notepadWindow.EditControl()
# Get the ValuePattern supported by the EditControl and set the control's text to "Hello" using the pattern.
edit.GetValuePattern().SetValue('Hello')  # or edit.GetPattern(auto.PatternId.ValuePattern)
edit.SendKeys('{Ctrl}{End}{Enter}World')  # Type at the end of the text
# First, find the TitleBarControl from the first layer of the notepadWindow control,
# and then find the second ButtonControl (the maximize button) among the descendants of the TitleBarControl, and click the button.
notepadWindow.TitleBarControl(Depth=1).ButtonControl(foundIndex=2).Click()
# Find the button with the name '关闭' (close) among the first two layers of descendants of the notepadWindow control and click the button.
notepadWindow.ButtonControl(searchDepth=2, Name='关闭').Click()
# At this point, a prompt asking whether to save the Notepad document pops up. Press Alt+N to exit without saving.
auto.SendKeys('{Alt}n')
```

The `auto.GetRootControl()` function returns the root node (desktop window) of the control tree. The `auto.WindowControl(searchDepth=1, ClassName='Notepad')` creates a `WindowControl` object, and the parameters in parentheses specify the conditions or control properties to search for this control in the control tree.

The `__init__` function of the `Control` class has the following parameters that can be used:

```
searchFromControl = None  # Where to start the search from. If None, start from the root node (Desktop).
searchDepth = 0xFFFFFFFF  # Search depth.
searchInterval = SEARCH_INTERVAL  # Search interval.
foundIndex = 1  # The index of the found control that satisfies the search conditions. Indices start from 1.
Name  # Control name.
SubName  # Partial control name.
RegexName  # Match control names using regular expressions. Only one of Name, SubName, and RegexName can be used at a time.
ClassName  # Control class name.
AutomationId  # Control AutomationId.
ControlType  # Control type.
Depth  # The exact depth of the control relative to searchFromControl.
Compare  # Custom comparison function function(control: Control, depth: int) -> bool.
```

The difference between `searchDepth` and `Depth` is as follows:
- `searchDepth` searches for the first control that satisfies the search conditions within the specified depth range (including all descendants in layers 1 to `searchDepth`).
- `Depth` searches for the first control that satisfies the search conditions only at the specified depth (excluding all descendants in layers 1 to `searchDepth-1`).

The `Control.Element` returns the underlying COM object [IUIAutomationElement](https://docs.microsoft.com/en-us/windows/win32/api/uiautomationclient/nn-uiautomationclient-iuiautomationelement) of the IUIAutomation COM interface. Essentially, all the properties and methods of `Control` are implemented by calling the IUIAutomationElement COM API and Win32 API. When you use a property or method of a `Control`, it internally calls `Control.Element`, and if `Control.Element` is `None

`, uiautomation starts searching for the control. If the control is not found within `uiautomation.TIME_OUT_SECOND` (default is 10 seconds), uiautomation raises a LookupError exception. Once a control is found, `Control.Element` will have a valid value. You can use `Control.Exists(maxSearchSeconds, searchIntervalSeconds)` to check if a control exists without throwing an exception. You can also use `Control.Refind` or `Control.Exists` to invalidate `Control.Element` and trigger a re-search.

Here's an example:

```python
#!python3
# -*- coding:utf-8 -*-
import subprocess
import uiautomation as auto
auto.uiautomation.SetGlobalSearchTimeout(15)  # Set the global search timeout to 15


def main():
    subprocess.Popen('notepad.exe')
    window = auto.WindowControl(searchDepth=1, ClassName='Notepad')
    edit = window.EditControl()
    # When SendKeys is called for the first time, uiautomation starts searching for the window and edit controls within 15 seconds
    # because SendKeys indirectly calls Control.Element, and Control.Element is None.
    # If the window and edit controls are not found within 15 seconds, a LookupError exception is raised.
    try:
        edit.SendKeys('first notepad')
    except LookupError as ex:
        print("The first notepad doesn't exist in 15 seconds")
        return
    # The second SendKeys call does not trigger a search because the previous call ensured that Control.Element is valid.
    edit.SendKeys('{Ctrl}a{Del}')
    window.GetWindowPattern().Close()  # Close the first Notepad, even though window and edit have a value, they are invalid.

    subprocess.Popen('notepad.exe')  # Run the second Notepad
    window.Refind()  # Must perform a new search
    edit.Refind()  # Must perform a new search
    edit.SendKeys('second notepad')
    edit.SendKeys('{Ctrl}a{Del}')
    window.GetWindowPattern().Close()  # Close the second Notepad, even though window and edit have a value, they are invalid again.

    subprocess.Popen('notepad.exe')  # Run the third Notepad
    if window.Exists(3, 1): # Trigger a new search
        if edit.Exists(3):  # Trigger a new search
            edit.SendKeys('third notepad')  # The previous Exists call ensures that edit.Element is valid
            edit.SendKeys('{Ctrl}a{Del}')
        window.GetWindowPattern().Close()
    else:
        print("The third notepad doesn't exist in 3 seconds")


if __name__ == '__main__':
    main()
```

You can also set `auto.uiautomation.DEBUG_SEARCH_TIME = True` to see the number of controls traversed and the search time when searching for controls.

Reference: demos/automation_calculator.py

## Demo

The **demos** directory provides some examples for learning how to use uiautomation.

## Q&A

### Cannot Find Controls

If you find that `automation.py` cannot print the controls of the program you are looking at, it is not a bug in uiautomation. This is because the program is implemented using DirectUI or custom controls, rather than the standard controls provided by Microsoft. This software must implement [UI Automation Provider](https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-providersoverview) in order to support UIAutomation. The standard controls provided by Microsoft typically support UIAutomation by default.

For example, in the Chrome browser, by default, you can only see the top-level `PaneControl` `Chrome_WidgetWin_1`, and you cannot see the specific child controls of Chrome. However, if you run the Chrome browser with the `--force-renderer-accessibility` parameter, you will be able to see the child controls of Chrome. This is because Chrome has implemented the UI Automation Provider and added a parameter switch. If a software is implemented using DirectUI but does not implement the UI Automation Provider, then it does not support UIAutomation.

## Some Screenshots

Batch rename PDF bookmarks
![bookmark](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\rename_pdf_bookmark.gif)

Get text from Microsoft Word
![Word](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\word.png)

Wireshark 3.0 (Qt 5.12)
![Wireshark](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\wireshark3.0.gif)

GitHub Desktop (Electron App)
![GitHubDesktop](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\github_desktop.png)

Display QQ
![QQ](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\automation_qq.png)

Print a nicely formatted directory structure
![PrettyPrint](C:\Users\yvzzi\Desktop\Python-UIAutomation-for-Windows\assets\readme.assets\pretty_print_dir.png)

.
