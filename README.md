# vbs-oofautomate
VBScript for automate the user's Out of Office state in Outlook using Exchange Web Services with autologon.

#### Set some variables:

Edit **ewsOOF.vbs**:
<pre>
<code>
userEmail = "mariusz@example.pl"
OofState = "Enabled"                     'Disabled or Enabled
ExternalAudience = "None"                'None or Known or All
InternalReply = "Out of office message"  'Internal message
ExternalReply = ""                       'External message
url = "https://exchangehost/EWS/Exchange.asmx" 'EWS Url
</code>
</pre>

#### Create sheduled task:

Run **CreateOOFScheduler.bat**. Script creates scheduled task:
<pre>
<code>
set "scriptPath="wscript //b //nologo '%cd%\ewsOOF.vbs'""
schtasks /create /tn OutOfOffice_Enable /tr %scriptPath% /sc weekly /d MON,TUE,WED,THU,FRI /st 15:55:00
</code>
</pre>
#### Delete sheduled task:

Run **DeleteOOFScheduler.bat**:
<pre>
<code>
schtasks /delete /tn OutOfOffice_Enable /f
</code>
</pre>
