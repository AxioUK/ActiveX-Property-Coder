Attribute VB_Name = "Module1"
Public currchar
Public thisLine
Public tstart
Public tend
Public holdtend
Public holdtstart
Public TopLine
Public foundpos
Public commentchar As String
Public longvar As String

Public Declare Function LockWindowUpdate Lib "user32" _
           (ByVal hwndLock As Long) As Long
           
           

Public Function ColorizeWord(Rich1 As RichTextBox, Word As String, color As OLE_COLOR)

      Do Until Rich1.GetLineFromChar(tstart) <> thisLine
            tstart = tstart - 1
            If tstart < 0 Then
                tstart = 0
                Exit Do
            End If
        
        Loop

startline = Rich1.GetLineFromChar(Rich1.SelStart)
If Rich1.SelLength > 0 Then Exit Function
Rich1.Enabled = False

tstart = tstart
If tstart = 0 Then
tstart = 1
End If


tstart = tstart - Len(Word)


Do
nowline = Rich1.GetLineFromChar(Rich1.SelStart)
If nowline <> startline Then GoTo endx
holdtstart = tstart + Len(Word)
commentposx = InStr(holdtstart, Rich1.Text, commentchar, vbTextCompare)
If holdtstart < 1 Then
holdtstart = 1
End If

tstart = tstart + Len(Word)
foundpos = InStr(tstart, Rich1.Text, Word, vbTextCompare)
If foundpos > tend Then GoTo endx '''''''''''''''''''''
If foundpos < 1 Then GoTo endx
If foundpos < 2 Then
sletter = ""
Else
sletter = Mid(Rich1.Text, foundpos - 1, 1)
End If
eletter = Mid(Rich1.Text, foundpos + Len(Word), 1)
If foundpos > 0 Then
If foundpos = 1 Then
tstart = tstart - 1
End If

Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
'###################################################
If Word = commentchar Then
 tend = Rich1.SelStart
       Do Until Rich1.GetLineFromChar(tend) <> thisLine
            tend = tend + 1
            If tend > Len(Rich1.Text) Then
                tend = Len(Rich1.Text) + 1
                Exit Do
            End If
        Loop
Rich1.SelStart = foundpos - 1
Rich1.SelLength = tend - (foundpos - 1)
Rich1.SelColor = color
Rich1.SelLength = 0
Rich1.SelStart = currchar
Rich1.SelColor = &H0&
Exit Function
Exit Do
End If
''''''''''''''''''''''''''''''
If Word = longvar Then
 tend = Rich1.SelStart
       Do Until Rich1.GetLineFromChar(tend) <> thisLine
            tend = tend + 1
            If tend > Len(Rich1.Text) Then
                tend = Len(Rich1.Text) + 1
                Exit Do
            End If
        Loop

pos = tstart
Do
foundpos = InStr(pos, Rich1.Text, longvar, vbTextCompare)

For i = foundpos To tend
If foundpos < 1 Then Exit For
If i = tend Then Exit For
Rich1.SelStart = i - 1
Rich1.SelLength = 1
If Rich1.SelText = "" Then Exit For
Select Case Asc(Rich1.SelText)
Case 48 To 57
Rich1.SelColor = color
Case 36
Rich1.SelColor = color
Case 97 To 122
Rich1.SelColor = color
Case 65 To 90
Rich1.SelColor = color
Case 145
Rich1.SelColor = color
Case 146
Rich1.SelColor = color
Case 143
Rich1.SelColor = color
Case 143
Rich1.SelColor = color
Case Else
Exit For
End Select

Next

pos = foundpos + 2
Loop While foundpos > 0

GoTo endx
End If


If tstart = 0 Then
tstart = 1
End If
commentposx = InStr(tstart, Rich1.Text, commentchar, vbTextCompare)
If commentposx > 0 Then
If Rich1.SelStart > commentposx Then GoTo endx
End If

If Len(Word) = 1 Then
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If



If eletter = "" And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If
If eletter = "" And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If
If eletter = " " And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If
If eletter = " " And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If
If eletter = "" And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

If eletter = " " And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

If eletter = Chr(10) And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If


If eletter = Chr(10) And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If


If eletter = Chr(10) And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

If eletter = Chr(13) And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

If eletter = Chr(13) And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

If eletter = Chr(13) And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = foundpos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = color
End If

Rich1.SelLength = 0
End If
If foundpos = 1 Then
tstart = tstart + 1
End If

Loop While foundpos > 0
endx:



Rich1.SelStart = currchar
Rich1.SelColor = &H0&
foundpos = 0
eletter = ""
sletter = ""
Rich1.Enabled = True

End Function

Public Function clearwordcolors(Rich1 As RichTextBox)

If Rich1.SelLength > 0 Then Exit Function
Rich1.Enabled = False
currchar = Rich1.SelStart

thisLine = Rich1.GetLineFromChar(Rich1.SelStart)
'Form1.Caption = KeyCode
tstart = Rich1.SelStart
tend = Rich1.SelStart
With Rich1
      Do Until .GetLineFromChar(tstart) <> thisLine
            tstart = tstart - 1
            If tstart < 0 Then
                tstart = 0
                Exit Do
            End If
        
        Loop



       Do Until .GetLineFromChar(tend) <> thisLine
            tend = tend + 1
            If tend > Len(.Text) Then
                tend = Len(.Text) + 1
                Exit Do
            End If

Loop
End With
If tstart = 1 Then
tend = tend - 2
End If
If tstart > 1 Then
tstart = tstart + 1
tend = tend - 1
End If
holdtstart = tstart
holdtend = tend
Rich1.SelStart = tstart
Rich1.SelLength = tend - tstart
Rich1.SelColor = &H0&
Rich1.SelLength = 0
Rich1.SelStart = currchar
holdtend = tend
Rich1.Enabled = True

End Function

