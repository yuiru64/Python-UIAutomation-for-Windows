@echo off
cl /c UIEventHandler.cpp
link -dll -out:UIEventHandler.dll .\UIEventHandler.obj
del .\UIEventHandler.obj