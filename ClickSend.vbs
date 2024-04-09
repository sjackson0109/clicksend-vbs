' This loops through a queue folder of TXT files, and sends the content to the configured API endpoint.
' [TXT to HTTP to SMS]
'
' See ClickSend API docs:
' https://developers.clicksend.com/docs/rest/v3

Dim API_BASE, API_ENDPOINT_SEND_SMS, QUEUE_PATH, INI_FILE, LOGSUCCESS

' GLOBAL CONFIG

LOGSUCCESS = True
INI_FILE = ".\ClickSend.ini"

' Initialise functions
Function ReadIni(myFilePath, mySection, myKey)
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    ReadIni = ""
    strFilePath = Trim(myFilePath)
    strSection = Trim(mySection)
    strKey = Trim(myKey)

    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim(objIniFile.ReadLine)
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                strLine = Trim(objIniFile.ReadLine)
                Do While Left(strLine, 1) <> "["
                    intEqualPos = InStr(1, strLine, "=", 1)
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        If LCase(strLeftString) = LCase(strKey) Then
                            ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            Exit Do
                        End If
                    End If
                    If objIniFile.AtEndOfStream Then Exit Do
                    strLine = Trim(objIniFile.ReadLine)
                Loop
                Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exist. Exiting..."
        Wscript.Quit 1
    End If
End Function

Function EncodeBase64(input)
    Dim objXML, objNode
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.CreateElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = StreamStringToBinary(input)
    EncodeBase64 = objNode.Text
End Function

Function StreamStringToBinary(str)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 'adTypeText
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText str
    objStream.Position = 0
    objStream.Type = 1 'adTypeBinary
    StreamStringToBinary = objStream.Read
    objStream.Close
    Set objStream = Nothing
End Function


' Extract Parameters
API_BASE = ReadIni(INI_FILE, "clicksend", "API_BASE")
If API_BASE = "" Then API_BASE = "."

API_USERNAME = ReadIni(INI_FILE, "clicksend", "API_USERNAME")
If API_USERNAME = "" Then API_USERNAME = "nocredit"

API_KEY = ReadIni(INI_FILE, "clicksend", "API_KEY")
If API_KEY = "" Then API_KEY = "D83DED51-9E35-4D42-9BB9-0E34B7CA85AE"

API_ENDPOINT_SEND_SMS = ReadIni(INI_FILE, "clicksend", "QueuePath")
If API_ENDPOINT_SEND_SMS = "" Then API_ENDPOINT_SEND_SMS = "/sms/send"

QUEUE_PATH = ReadIni(INI_FILE, "clicksend", "QueuePath")
If QUEUE_PATH = "" Then QUEUE_PATH = ".\Queue\"

' Begin

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(QUEUE_PATH)
Set colFiles = objFolder.Files 'Queue of files

WScript.Echo "Starting to process files in the queue folder..."

For Each objFile In colFiles
    Dim EvenFileName, EvenFileExt, EvenFile
    EvenFileName = objFile.Name
    EvenFileExt = objFSO.GetExtensionName(EvenFileName)
    EvenFile = QUEUE_PATH & EvenFileName
    If EvenFileExt = "txt" Or EvenFileExt = "log" Or EvenFileExt = "json" Then
        WScript.Echo "Processing file: " & EvenFile
        Dim POST_BODY, LockingFile, EvenFileContent, Status, Response
        LockingFile = EvenFile & ".lock"
        If objFSO.FileExists(LockingFile) Then
            WScript.Echo "Lock file already exists: " & LockingFile & ". Moving on to the next file."
        Else
            WScript.Echo "Creating lock file: " & LockingFile
            objFSO.CreateTextFile(LockingFile)
            Set EvenFileContent = CreateObject("ADODB.Stream")
            EvenFileContent.CharSet = "utf-8"
            EvenFileContent.Open
            EvenFileContent.LoadFromFile(EvenFile)
            POST_BODY = EvenFileContent.ReadText()
            POST_BODY = Replace(POST_BODY, "\", "\\")
            POST_BODY = Replace(POST_BODY, vbCr, "")
            POST_BODY = Replace(POST_BODY, vbLf, "")
            EvenFileContent.Close
            ' Send the alert file content to ClickSend and check the response
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            objHTTP.Open "POST", API_BASE & API_ENDPOINT_SEND_SMS, False
            objHTTP.SetRequestHeader "Authorization", "Basic " & EncodeBase64(API_USERNAME & ":" & API_KEY)
            objHTTP.SetRequestHeader "Content-Type", "application/json"
            objHTTP.Send POST_BODY
            If objHTTP.Status = 200 Then
                objFSO.DeleteFile(EvenFile)
                WScript.Echo "File sent successfully and deleted: " & EvenFile
            Else
                WScript.Echo "Error from the remote webserver"
            End If
            WScript.Echo "Deleting lock file: " & LockingFile
            objFSO.DeleteFile(LockingFile)
        End If

    Else
        ' SKIP ALL OTHER NON-SUPPORTED FILE
    End If
Next

WScript.Echo "Processing of files in the queue folder completed."