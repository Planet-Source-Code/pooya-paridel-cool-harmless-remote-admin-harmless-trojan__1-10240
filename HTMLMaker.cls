VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function MakeFormPage(Title As String, Action As String, Ttext As String, InputData As Collection) As String

    Dim ipos As Integer
    Dim jpos As Integer
    
    Dim tstr As String
        
    tstr = tstr + "<html>" & vbCrLf
    tstr = tstr + "<head>" & vbCrLf
    tstr = tstr + "<title>" & Title & "</title>" & vbCrLf
    tstr = tstr + "</head>" & vbCrLf
    tstr = tstr + "<body>" & vbCrLf
    tstr = tstr + "<p align='center'><form action='/' method=GET>" & vbCrLf
    tstr = tstr + "<table bgcolor=#FFFFF4 border=2 width=85%  cellspacing=0><tr><td>"
    tstr = tstr + "<center><table border=0>" & vbCrLf
    tstr = tstr + "<tr><td colspan=2><Font face='Homa'>" & Ttext & vbCrLf
    tstr = tstr + "<input Type='Hidden' Name='Action'  Value='" & Action & "'></td></tr>" & vbCrLf
    
    For ipos = 1 To InputData.Count Step 5

        If UCase(InputData(ipos + 1)) = "SELECT" Then
            
                tstr = tstr + "<tr><td><Font face='Homa'>" & InputData(ipos) & "</td><td><select name='" & InputData(ipos + 2) & "' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">" & InputData(ipos + 3) & "</select></td></tr>" & vbCrLf
        
        ElseIf UCase(InputData(ipos + 1)) = "DATE" Then
                
                tstr = tstr + "<tr><td><Font face='Homa'>" & InputData(ipos) & "</td><td>" & vbCrLf
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "Y' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1998 To 2099
                    tstr = tstr + "<option>" & Str(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select>" & vbCrLf
                
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "M' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1 To 12
                    tstr = tstr + "<option>" & Str(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select>" & vbCrLf
                
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "D' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1 To 31
                    tstr = tstr + "<option>" & Str(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select></td></tr>" & vbCrLf
                
        ElseIf UCase(InputData(ipos + 1)) = "SDATE" Then
                
                tstr = tstr + "<tr><td><Font face='Homa'>" & InputData(ipos) & "</td><td>" & vbCrLf
                
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "D' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1 To 31
                    tstr = tstr + "<option>" & Str(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select>" & vbCrLf
                                
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "M' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1 To 12
                    tstr = tstr + "<option value=" & CStr(jpos) & ">" & PFun.MonthName(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select>" & vbCrLf
                                
                tstr = tstr + "<select name='" & InputData(ipos + 2) & "Y' size=1 "
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">"
                For jpos = 1377 To 1420
                    tstr = tstr + "<option>" & Str(jpos) & "</option>" & vbCrLf
                Next
                tstr = tstr + "</select></td></tr>" & vbCrLf
                
        ElseIf UCase(InputData(ipos + 1)) = "HIDDEN" Then
                
                tstr = tstr + "<input Type='Hidden' "
                
                If InputData(ipos + 2) <> "" Then
                    tstr = tstr + "Name='" & InputData(ipos + 2) & "' "
                End If
                If InputData(ipos + 3) <> "" Then
                    tstr = tstr + "Value='" & InputData(ipos + 3) & "' "
                End If
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + ">" & vbCrLf
        
        ElseIf UCase(InputData(ipos + 1)) = "CHECKBOX" Then
        
                tstr = tstr + "<tr><td colspan=2><input Type='Checkbox' "
                
                If InputData(ipos + 2) <> "" Then
                    tstr = tstr + "Name='" & InputData(ipos + 2) & "' "
                End If
                If InputData(ipos + 3) <> "" Then
                    tstr = tstr + InputData(ipos + 3)
                End If
                
                tstr = tstr + "><font face='Homa'>" & InputData(ipos) & "</td></tr>" & vbCrLf
    
        ElseIf UCase(InputData(ipos + 1)) = "TEXTAREA" Then
        
                If InputData(ipos) = "" Then
                    tstr = tstr + "<tr><td colspan=2><textarea rows='5' cols='30' "
                Else
                    tstr = tstr + "<tr><td><Font face='Homa'>" & InputData(ipos) & "</td><td><textarea rows='5' cols='30' "
                End If

                tstr = tstr + "Name='" & InputData(ipos + 2) & "' >"

                If InputData(ipos + 3) <> "" Then
                    tstr = tstr + InputData(ipos + 3)
                End If
                tstr = tstr + "</textarea></td></tr>" & vbCrLf
                
        Else
        
                If InputData(ipos) = "" Then
                    tstr = tstr + "<tr><td colspan=2><input "
                Else
                    tstr = tstr + "<tr><td><Font face='Homa'>" & InputData(ipos) & "</td><td><input "
                End If
            
                If InputData(ipos + 1) <> "" Then
                    tstr = tstr + "Type='" & InputData(ipos + 1) & "' "
                End If
                If InputData(ipos + 2) <> "" Then
                    tstr = tstr + "Name='" & InputData(ipos + 2) & "' "
                End If
                If InputData(ipos + 3) <> "" Then
                    tstr = tstr + "Value='" & InputData(ipos + 3) & "' "
                End If
                If InputData(ipos + 4) <> "" Then
                    tstr = tstr + "Style='" & InputData(ipos + 4) & "' "
                End If
                tstr = tstr + "></td></tr>" & vbCrLf
        
        End If
    Next
    
    tstr = tstr + "</table></td></tr></table>" & vbCrLf
    tstr = tstr + "</p>"
    tstr = tstr + CopyRight
    tstr = tstr + "</body>" & vbCrLf
    tstr = tstr + vbCrLf
    tstr = tstr + "</html>" & vbCrLf

    MakeFormPage = tstr
End Function

Sub AddTo(Src As Collection, iText As String, iType As String, iName As String, iValue As String, iStyle As String)
        
        Src.Add iText: Src.Add iType
        Src.Add iName: Src.Add iValue
        Src.Add iStyle
        
End Sub

Public Function MakeMsgForm(Title As String, Msg As String) As String

        Dim tstr As String
        
        tstr = tstr + "<html>" & vbCrLf
        tstr = tstr + "<head>" & vbCrLf
        tstr = tstr + "<title>" & Title & "</title>" & vbCrLf
        tstr = tstr + "</head>" & vbCrLf
        tstr = tstr + "<body>" & vbCrLf
        tstr = tstr + "<p align='center'>" & vbCrLf
        tstr = tstr + "<table bgcolor=#FFF1F1 border=2 width=75%  cellspacing=0><tr><td><center><Font face='Homa'>" & vbCrLf
        tstr = tstr + "<br>" & Msg & "<br><br>" & vbCrLf
        tstr = tstr + "</font></td></tr></table>" & vbCrLf
        tstr = tstr + "</p>"
        tstr = tstr + CopyRight
        tstr = tstr + "</body>" & vbCrLf
        tstr = tstr + vbCrLf
        tstr = tstr + "</html>" & vbCrLf
        
        MakeMsgForm = tstr
End Function

Sub MakeQuestion(Dir As String, Persian As Boolean, Title As String, Msg As String, YUrl, NUrl)

        cgi.WriteCGI "<html dir=" & Dir & ">" & vbCrLf
        cgi.WriteCGI "<head>" & vbCrLf
        If Persian = True Then
            cgi.WriteCGI "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1256"">"
        End If
        cgi.WriteCGI "<title>" & Title & "</title>" & vbCrLf
        cgi.WriteCGI "<font face=Homa>" & vbCrLf
        cgi.WriteCGI "</head>" & vbCrLf
        cgi.WriteCGI "<body>" & vbCrLf
        cgi.WriteCGI "<p align='center'>" & vbCrLf
        cgi.WriteCGI "<table bgcolor=#F1F1FF border=2 width=75%  cellspacing=0>" & vbCrLf
        cgi.WriteCGI "<tr><td colspan=2><center><br>" & Msg & "<br></font></td></tr>" & vbCrLf
        cgi.WriteCGI "<tr><td><center><A href=" & NUrl & ">ÎíÑ</a></td><td><center><A href=" & YUrl & ">Èáå</a></td></tr>"
        cgi.WriteCGI "</table>" & vbCrLf
        cgi.WriteCGI "</p>"
        cgi.WriteCGI "</body>" & vbCrLf
        cgi.WriteCGI vbCrLf
        cgi.WriteCGI CopyRight
        cgi.WriteCGI "</html>" & vbCrLf

End Sub

Function CopyRight() As String
    CopyRight = "<br><p align=center><center><table bordercolor=#FFEEE1 bgcolor=#FFF5E9 border=1  cellspacing=0 cellpadding=0><tr><td><center><font face='Comic,Arial' size=-2>VB Tro v" & App.Major & "." & App.Minor & " r" & App.Revision & "<br>by Pooya Paridel<br>Dont Use It Like a Trojan</font></td></tr></table></p>"

End Function

