@echo off
title Khoi chay Reminder
echo Dang kiem tra va cai dat thu vien schedule...
python -m pip install schedule

echo.
echo Dang khoi chay ung dung o che do chay ngam...
:: Su dung start "" de chay lenh va dong cua so CMD ngay lap tuc
start "" pythonw.exe ".\reminder.pyw"

echo.
echo Da kich hoat thong bao thanh cong!
echo Cua so nay se tu dong dong sau 3 giay.
timeout /t 3 /nobreak > nul
exit