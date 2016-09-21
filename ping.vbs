Option Explicit

Dim strHost, strFile

strHost = "www.google.com" '"127.0.0.1"
strFile = "C:\Test.txt"

PingForever strHost, strFile

Sub PingForever(strHost, outputfile)
    Dim Output, Shell, strCommand, ReturnCode

    Set Output = CreateObject("Scripting.FileSystemObject").OpenTextFile(outputfile, 8, True)
    Set Shell = CreateObject("wscript.shell")
    strCommand = "ping -n 1 -w 300 " & strHost
    While(True)
        ReturnCode = Shell.Run(strCommand, 0, True)     
        If ReturnCode = 0 Then
            Output.WriteLine Date() & " - " & Time & " | O Servidor " & strHost & " esta online"
        Else
            Output.WriteLine Date() & " - " & Time & " | O Servidor " & strHost & " esta offline"
        End If
        Wscript.Sleep 2000
    Wend
End Sub
