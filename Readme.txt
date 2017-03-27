Files
AutoDeploy.txt  ---------- Log sample
SFC_Deploy.vbs   ---------- main program


2017-02-15 change log
1.新增 Windows7 .Net Framework & KB2468871 安裝
	1.1 自動判斷Windows 7 x64/x32
	1.2 依x64/x32 KB版本不同, 進行安裝
2.程式執行前會先進行Log清除, 只留最一次執行的Log


2017-02-15 change log
1. 硬碟空間不足1G,Fail。
2. IE8 最後安裝, 安裝完整請使用者重起電腦(Alert)。
3. Log檔名稱修改為AutoDeploy.txt


2017-01-20 change log
1. Added check .netframwork and KB2468871 install status even if IE8 installed
2. Added install status check(succeed/failed) after software install process terminated


2017-01-18 change log
1. The install process only support XP
2. Source download will save to C:\SFC_D
3. Source folder and files wil delete when process terminated
4. Log file will save to C:\SFCS\deploy_log
5. C:\SFCS\deploy_log folder always be created on every OS when process is running
