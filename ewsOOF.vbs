Option Explicit

Dim userEmail 'Email address
Dim OofState 'Disabled or Enabled
Dim ExternalAudience 'None or Known or All
Dim InternalReply '
Dim ExternalReply '
Dim ExchangeUrl

userEmail = "mariusz@example.pl"
OofState = "Enabled"
ExternalAudience = "None"
InternalReply = "Out of office message"
ExternalReply = ""
ExchangeUrl = "https://poczta/EWS/Exchange.asmx"
	
Public Function buildXml()

    Dim xmlDoc, objRoot, objBody, tmpAttr


    Set xmlDoc = CreateObject("Microsoft.XMLDOM")  

    Set objRoot = xmlDoc.createElement("soap:Envelope")
        tmpAttr = objRoot.setAttribute("xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance")
        tmpAttr = objRoot.setAttribute("xmlns:xsd","http://www.w3.org/2001/XMLSchema")
        tmpAttr = objRoot.setAttribute("xmlns:soap","http://schemas.xmlsoap.org/soap/envelope/")
        xmlDoc.appendChild objRoot

    Set objBody = xmlDoc.createElement("soap:Body")
        objRoot.appendChild objBody

    Dim objSetUserOofSettingsRequest
    Set objSetUserOofSettingsRequest = xmlDoc.createElement("SetUserOofSettingsRequest")
        tmpAttr = objSetUserOofSettingsRequest.setAttribute("xmlns","http://schemas.microsoft.com/exchange/services/2006/messages")
        objBody.appendChild objSetUserOofSettingsRequest

    Dim objMailbox
    Set objMailbox = xmlDoc.createElement("Mailbox")
        tmpAttr = objMailbox.setAttribute("xmlns","http://schemas.microsoft.com/exchange/services/2006/types")
    Dim objMailboxAddress
    Set objMailboxAddress = xmlDoc.createElement("Address")
        objMailboxAddress.Text = userEmail
        objMailbox.appendChild objMailboxAddress
        objSetUserOofSettingsRequest.appendChild objMailbox

    Dim objUserOofSettings
    Set objUserOofSettings = xmlDoc.createElement("UserOofSettings")
        tmpAttr = objUserOofSettings.setAttribute("xmlns","http://schemas.microsoft.com/exchange/services/2006/types")
        objSetUserOofSettingsRequest.appendChild objUserOofSettings

    Dim objOofState
    Set objOofState = xmlDoc.createElement("OofState")
        objOofState.Text = OofState
        objUserOofSettings.appendChild objOofState

    Dim objExternalAudience
    Set objExternalAudience = xmlDoc.createElement("ExternalAudience")
        objExternalAudience.Text = ExternalAudience
        objUserOofSettings.appendChild objExternalAudience


    Dim objInternalReply, objInternalReplyMessage
    Set objInternalReply = xmlDoc.createElement("InternalReply")
    Set objInternalReplyMessage = xmlDoc.createElement("Message")
        objInternalReplyMessage.Text = InternalReply
        objInternalReply.appendChild objInternalReplyMessage
        objUserOofSettings.appendChild objInternalReply

    Dim objExternalReply, objExternalReplyMessage
    Set objExternalReply = xmlDoc.createElement("ExternalReply")
    Set objExternalReplyMessage = xmlDoc.createElement("Message")
        objExternalReplyMessage.Text = ExternalReply
        objExternalReply.appendChild objExternalReplyMessage
        objUserOofSettings.appendChild objExternalReply

    buildXml = xmlDoc.xml

    Set objBody = Nothing
    Set objExternalAudience = Nothing
    Set objExternalReply = Nothing
    Set objExternalReplyMessage = Nothing
    Set objInternalReply = Nothing
    Set objInternalReplyMessage = Nothing
    Set objMailbox = Nothing
    Set objMailboxAddress = Nothing
    Set objOofState = Nothing
    Set objRoot = Nothing
    Set objSetUserOofSettingsRequest = Nothing
    Set objUserOofSettings = Nothing

End Function

Public Function sendRequest(url, method, data, contentType)
    Dim objWinHttp, Err
    On Error Resume Next
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    objWinHttp.SetAutoLogonPolicy 0
    objWinHttp.Open method, url
    objWinHttp.setRequestHeader "Content-type", contentType
    objWinHttp.Send(data)

    if Err.Number <> 0 then
        If objWinHttp.Status = "200" Then
            sendRequest = objWinHttp.ResponseText
        ElseIf objWinHttp.Status = "500" Then
            sendRequest = objWinHttp.ResponseText
        else
            sendRequest = "HTTP " & objWinHttp.Status & " " & _
            objWinHttp.StatusText & " RESPONSE: " & objWinHttp.ResponseText
        End If
    Else
        sendRequest = Err.Number & " SRC: " & Err.Source & " DST: " &  Err.Description
        Err.Clear
    end if
    On Error GoTo 0
    Set objWinHttp = Nothing

End Function

Public Function getXmlValueFromTagName(xmlString, tagName, nodeIndex)
    Dim doc
    Dim nodes
    Set doc = CreateObject("MSXML2.DOMDocument")
    doc.loadXML xmlString

    Set nodes = doc.selectNodes("//" & tagName)
    getXmlValueFromTagName = nodes(nodeIndex).text

    Set doc = Nothing
    Set nodes = Nothing

End Function

Dim xmlString, url, method, data, contentType

url = ExchangeUrl
method = "POST"
data = buildXml()
contentType = "text/xml; charset=utf-8"

Dim response
response = sendRequest(url, method, data, contentType)
On Error Resume Next
Dim responseCode
responseCode = getXmlValueFromTagName(response,"ResponseCode",0)
if Err.Number <> 0 Or responseCode <> "NoError"  then
    Dim errString
    if Err.Number <> 0 then
        errString = Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description & " "
    end if
    errString = errString & " Error: " & response

    Dim FSO, outFile
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set outFile = FSO.OpenTextFile("log.txt" ,8 , True)
    outFile.Write(errString)
    outFile.Close

    Set FSO = Nothing

    MsgBox "ewsOOF.vbs, check log.txt"

End If
On Error GoTo 0
MsgBox "Done!"
