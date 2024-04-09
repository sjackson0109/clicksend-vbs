' This loops through a queue folder of TXT files, and sends the content to the configured API endpoint.
' [TXT to HTTP to SMS]
'
' See ClickSend API docs:
' https://developers.clicksend.com/docs/rest/v3

Dim API_BASE, API_ENDPOINT_SEND_SMS, QUEUE_PATH, INI_FILE, LOGSUCCESS
' Constants and shell object used for logging to the Windows Application Event Log

Const EVENT_SUCCESS	= 0
Const EVENT_ERROR 	= 1
Const EVENT_WARNING = 2
Const EVENT_INFO 	= 4
Set objShell = WScript.CreateObject("WScript.Shell")

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
                            ' Remove any double quotes from the output
                            ReadIni = Replace(ReadIni, Chr(34), "")
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

Function Base64Encode(sText)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

' Extract Parameters
API_USERNAME = ReadIni(INI_FILE, "clicksend", "API_USERNAME")
If API_USERNAME = "" Then API_USERNAME = "nocredit"

API_KEY = ReadIni(INI_FILE, "clicksend", "API_KEY")
If API_KEY = "" Then API_KEY = "D83DED51-9E35-4D42-9BB9-0E34B7CA85AE"

API_ENDPOINT = ReadIni(INI_FILE, "clicksend", "API_ENDPOINT")
If API_ENDPOINT = "" Then API_ENDPOINT = "https://rest.clicksend.com/v3/sms/send"

QUEUE_PATH = ReadIni(INI_FILE, "clicksend", "QueuePath")
If QUEUE_PATH = "" Then QUEUE_PATH = ".\Queue\"

' Begin
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(QUEUE_PATH)
Set colFiles = objFolder.Files 'Queue of files

WScript.Echo "Inspecting Queue folder"
For Each objFile In colFiles
    Dim EvenFileName, EvenFileExt, EvenFile
    EvenFileName = objFile.Name
    EvenFileExt = objFSO.GetExtensionName(EvenFileName)
    EvenFile = QUEUE_PATH & EvenFileName
    If EvenFileExt = "txt" Or EvenFileExt = "log" Or EvenFileExt = "json" Then
        WScript.Echo " Processing: " & EvenFile
        Dim POST_BODY, LockingFile, EvenFileContent, Status, Response
        LockingFile = EvenFile & ".lock"
        If objFSO.FileExists(LockingFile) Then
            WScript.Echo "  LOCK file found:" & LockingFile & ". Skipping."
        Else
            WScript.Echo "  LOCK file created: " & LockingFile
            objFSO.CreateTextFile(LockingFile)
            Set EvenFileContent = CreateObject("ADODB.Stream")
            EvenFileContent.CharSet = "utf-8"
            EvenFileContent.Open
            EvenFileContent.LoadFromFile(EvenFile)
            POST_BODY = Replace(Replace(Replace(EvenFileContent.ReadText(), "\", "\\"), vbCr, ""), vbLf, "")
            EvenFileContent.Close
            ' Send the message content to ClickSend and check the response
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            objHTTP.SetTimeouts 30,500,1000,1000
            objHTTP.Open "POST", API_ENDPOINT, False', API_USERNAME, API_KEY
            objHTTP.SetRequestHeader "Accept-Chartset", "utf-8"
            objHTTP.SetRequestHeader "Content-Type", "application/json"
                        WScript.Echo Base64Encode(API_USERNAME)
            objHTTP.SetRequestHeader "Authorization", "Basic " & Base64Encode(API_USERNAME & ":" & API_KEY)
            objHTTP.Send POST_BODY   'POST SMS
            If Err.Number <> 0 Then
                WScript.Echo "  ERROR: Couldn't connect or send data to Remote Server. Check Windows Application Event Log for details."

                objShell.LogEvent EVENT_ERROR, "Couldn't connect or send data to Remote Server." & vbNewLine & vbNewLine &_
                    "Error Number: " & Err.Number & vbNewLine & "Source: " & Err.Source & vbNewLine & "Description: " & Err.Description
            Else
                Status = objHTTP.Status
                Response = objHTTP.responseText
                WScript.Echo "  Response from Remote Server is: [" & Status & "] " & Response

                ' Remove the alert from the queue once it has been accepted by Remote Server,
                ' or log why the event wasn't accepted by Remote Server.

                If Status = 200 Then
                    WScript.Echo "  Deleting message: " & EvenFile
                    objFSO.DeleteFile(EvenFile)
                    
                    WScript.Echo "  Deleting lock file: " & LockingFile
                    objFSO.DeleteFile(LockingFile)

                    If LogSuccess = True Then
                        objShell.LogEvent EVENT_SUCCESS, "  Remote Server accepted event with data:" & vbNewLine & vbNewLine &_
                            PostBody & vbNewLine & vbNewLine & "Response was:" & vbNewLine & vbNewLine & "[" & Status & "] " & Response
                    End If
                Else
                    WScript.Echo "  Non-200 response received. Keeping message in queue: " & EvenFile

                    objShell.LogEvent EVENT_ERROR, "Remote Server did not accept event with data:" & vbNewLine & vbNewLine &_
                        PostBody & vbNewLine & vbNewLine & "Response was:" & vbNewLine & vbNewLine & "[" & Status & "] " & Response
                    
                    WScript.Echo "  Deleting lock file: " & LockingFile
                    objFSO.DeleteFile(LockingFile)
                End If
            End If
        End If
        WScript.Echo " Next file..."
    Else
        ' SKIP ALL OTHER NON-SUPPORTED FILE
    End If
Next

WScript.Echo "Processing of files in the queue folder completed."