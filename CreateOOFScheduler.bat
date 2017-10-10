set "scriptPath="wscript //b //nologo '%cd%\ewsOOF.vbs'""
schtasks /create /tn OutOfOffice_Enable /tr %scriptPath% /sc weekly /d MON,TUE,WED,THU,FRI /st 15:55:00