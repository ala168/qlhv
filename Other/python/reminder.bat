@echo off
title Khoi chay Reminder
echo Dang kiem tra va cai dat thu vien schedule...
python -m pip install schedule

start "" pythonw.exe ".\reminder.pyw"

start "" javaw -jar ".\wording.jar"

exit