Attribute VB_Name = "WebFunc"
Option Explicit
Option Compare Text
Option Base 1

'Web Requests Function Library
'Version 1.1.1

'Imports
'Microsoft WinHTTP Services, version 5.1
'Microsoft Scripting Runtime Reference
'Microsoft ActiveX Data Objects 6.1 Library

'History
' 1.0.0 - Initial Verion
' 1.0.1 - Added SendSQLCommand
' 1.1.0 - Added UploadTextViaFTP
' 1.1.1 - Removed MsgBox after FTP Upload

Public Function MakeWebRequest(ByVal Method As String, ByVal URL As String, PostData) As String
    ' make sure to include the Microsoft WinHTTP Services in the project
    ' tools -> references -> Microsoft WinHTTP Services, version 5.1
    ' http://www.808.dk/?code-simplewinhttprequest
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384106(v=vs.85).aspx
    ' http://www.neilstuff.com/winhttp/

    ' create the request object
    Dim Request As New WinHttpRequest

    ' set timeouts
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384061(v=vs.85).aspx
    ' SetTimeouts(resolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout)
    Request.SetTimeouts 60000, 60000, 60000, 60000

    ' make the request, http verb (method), url, false to force syncronous
    ' open(http method, absolute uri to request, async (true: async, false: sync)
    Request.Open Method, URL, False

    ' handle post content type
    If Method = "POST" Then
        Request.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End If

    ' set WinHttpRequestOption enumerations
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384108(v=vs.85).aspx

    ' set user agent
    Request.Option(0) = "Echovoice VBA HTTP Bot v0.1"

    ' set ssl ignore errors
    '   13056: ignore errors
    '   0: break on errors
    Request.Option(4) = 13056

    ' set redirects
    Request.Option(6) = True

    ' allow http to redirect to https
    Request.Option(12) = True

    ' send request
    ' send post data, should be blank for a get request
    Request.Send PostData

    ' read response and return
    MakeWebRequest = Request.ResponseText

End Function

Public Function ParseJSON(data As String) As Object
    ' http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
    ' take JSON and convert to object

    ' change code in cStringBuilder for Win64 systems
    ' https://code.google.com/p/vba-json/issues/detail?id=13

    ' add Microsoft Scripting Runtime Reference for Dictionary data type
    ' add ADO reference Microsoft ActiveX Data Objects 2.8 Library
    ' http://msdn.microsoft.com/en-us/library/aa241766(v=vs.60).aspx

    ' use Set when returning an object
    ' Set ParseJSON = Parse(data)

End Function

Public Sub SendSQLCommand(server As String, database As String, sqlQuery As String, responseCell As String)

  Dim connectionString As String
  Dim connectObj As ADODB.Connection
  Dim commandObj As ADODB.Command
  Dim rowsAffected As Long

  connectionString = "Provider=MSOLEDBSQL;" & _
                     "Server=" & server & ";" & _
                     "Database=" & database & ";" & _
                     "Integrated Security=SSPI;" & _
                     "TrustServerCertificate=Yes;" & _
                     "Encrypt=No;"

  On Error GoTo ErrorHandler

  Set connectObj = New ADODB.Connection
  connectObj.Open connectionString
  Set commandObj = New ADODB.Command
  With commandObj
    .ActiveConnection = connectObj
    .CommandText = sqlQuery
    .CommandType = adCmdText
    .Execute rowsAffected
  End With
  connectObj.Close
  Set commandObj = Nothing
  Set connectObj = Nothing

  Range(responseCell).Value = rowsAffected & " Rows Affected"

Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  If Not connectObj Is Nothing Then
    If connectObj.State = adStateOpen Then connectObj.Close
  End If
  Set commandObj = Nothing
  Set connectObj = Nothing

  Range(responseCell).Value = "Error sending data"

End Sub

Public Sub UploadTextViaFTP(server As String, user As String, pass As String, file As String, content As String)
    Dim TempFile As String, FTPCommandFile As String
    Dim fNum As Integer

    ' Create temporary text file
    TempFile = Environ$("TEMP") & "\temp_upload.txt"
    fNum = FreeFile
    Open TempFile For Output As #fNum
    Print #fNum, content
    Close #fNum

    ' Create temporary FTP command file
    FTPCommandFile = Environ$("TEMP") & "\ftp_commands.txt"
    fNum = FreeFile
    Open FTPCommandFile For Output As #fNum
    Print #fNum, "open " & server
    Print #fNum, user
    Print #fNum, pass
    Print #fNum, "binary"
    Print #fNum, "put " & TempFile & " " & file
    Print #fNum, "bye"
    Close #fNum

    ' Run FTP command
    Shell "cmd.exe /c ftp -s:""" & FTPCommandFile & """", vbNormalFocus

    ' Optional: wait a bit then clean up
    Application.Wait Now + TimeValue("0:00:05")
    Kill TempFile
    Kill FTPCommandFile
End Sub
