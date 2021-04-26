@echo off
::echo The current directory is [%cd%]
::copy app.py app_old.py /-Y /V
::C:\Users\drift\AppData\Local\Programs\Python\Python39\Scripts\pyuic5.exe -x app.ui -o app.py
::pause

::=========================== Get date and time ====================
::set CUR_YYYY=%date:~6,4%
::set CUR_MM=%date:~3,2%
::set CUR_DD=%date:~0,2%

::set CUR_HH=%time:~0,2%
::set CUR_NN=%time:~3,2%
::set CUR_SS=%time:~6,2%
::set CUR_MS=%time:~9,2%

::set SUBFILENAME=%CUR_YYYY%%CUR_MM%%CUR_DD%-%CUR_HH%%CUR_NN%%CUR_SS%
::set SUBFILENAME=%CUR_HH%%CUR_MM%%CUR_SS%-%CUR_DD

::echo The current date and time is [%SUBFILENAME%]

::========== Create .py file from .ui file =========================
::C:\Users\drift\AppData\Local\Programs\Python\Python39\Scripts\pyuic5.exe -x "app.ui" -o "app_%SUBFILENAME%.py"