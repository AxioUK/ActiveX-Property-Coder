VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ActiveX Coder 4"
   ClientHeight    =   7725
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRefresh 
      Caption         =   "Refresh on LET?"
      Height          =   270
      Left            =   5460
      TabIndex        =   21
      Top             =   1395
      Width           =   1635
   End
   Begin VB.OptionButton OptDIM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Private"
      Height          =   285
      Index           =   1
      Left            =   5460
      TabIndex        =   20
      Top             =   375
      Value           =   -1  'True
      Width           =   1110
   End
   Begin VB.OptionButton OptDIM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dim"
      Height          =   285
      Index           =   0
      Left            =   5460
      TabIndex        =   19
      Top             =   135
      Width           =   1110
   End
   Begin VB.CheckBox chkReduced 
      Caption         =   "Reduced code"
      Height          =   270
      Left            =   5460
      TabIndex        =   18
      Top             =   1155
      Width           =   1635
   End
   Begin ComctlLib.ListView lstmain 
      Height          =   2325
      Left            =   75
      TabIndex        =   6
      Top             =   1680
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      SmallIcons      =   "imglst1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4230
      TabIndex        =   17
      Top             =   1380
      Width           =   1110
   End
   Begin VB.CommandButton cmdInvertSel 
      Caption         =   "Invert Sel."
      Height          =   255
      Left            =   2032
      TabIndex        =   16
      Top             =   1380
      Width           =   990
   End
   Begin VB.CommandButton cmdUnSel 
      Caption         =   "UnSel All"
      Height          =   255
      Left            =   1016
      TabIndex        =   15
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Sel All"
      Height          =   255
      Left            =   90
      TabIndex        =   14
      Top             =   1380
      Width           =   885
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmMain.frx":1D82
      Left            =   2835
      List            =   "frmMain.frx":1D84
      TabIndex        =   1
      Top             =   270
      Width           =   2445
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7290
      TabIndex        =   7
      ToolTipText     =   "Copy generated code"
      Top             =   1140
      Width           =   1500
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   2445
   End
   Begin VB.TextBox txtDefValue 
      Height          =   315
      Left            =   2835
      TabIndex        =   3
      Top             =   855
      Width           =   2445
   End
   Begin VB.TextBox txtVariable 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   2445
   End
   Begin VB.CommandButton cmdAddList 
      Caption         =   "Add"
      Height          =   375
      Left            =   5445
      TabIndex        =   4
      ToolTipText     =   "Add values"
      Top             =   705
      Width           =   1140
   End
   Begin VB.CommandButton cmdRemoveSel 
      Caption         =   "Remove Sel."
      Height          =   255
      Left            =   3063
      TabIndex        =   5
      ToolTipText     =   "Remove selected entrie"
      Top             =   1380
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   5355
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rCode 
      Height          =   3510
      Left            =   75
      TabIndex        =   13
      Top             =   4080
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   6191
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":1D86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   6315
      Picture         =   "frmMain.frx":1E02
      ScaleHeight     =   1200
      ScaleWidth      =   2580
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   -120
      Width           =   2580
   End
   Begin VB.Label lbllstcount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   -330
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of variable:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   2835
      TabIndex        =   10
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Container variable:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   660
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Value:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2865
      TabIndex        =   8
      Top             =   660
      Width           =   1080
   End
   Begin ComctlLib.ImageList imgSmall 
      Left            =   4665
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":64D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6D74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImp 
         Caption         =   "Import..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Co&py"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "A&dd"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "Re&move Selected"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRemoveAll 
         Caption         =   "Remove &All..."
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGenCode 
         Caption         =   "Generate &Code"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuLstMain 
      Caption         =   "lstmain"
      Visible         =   0   'False
      Begin VB.Menu mnuLstMainEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuLstMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLstMainRemove 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuLstMainRemoveAll 
         Caption         =   "Remove &All..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tHt As LVHITTESTINFO
Private list_item As ListItem
Private X As New clsList
Private i As Integer
Private DocChanged As Boolean
Private docname As String
Private xx As Integer

Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Dim TT As CTooltip
Dim m_lCurItemIndex As Long

Public Function Capitalize(ByVal pstrText As String) As String
If Left$(pstrText, 1) <> UCase$(Left$(pstrText, 1)) Then
  pstrText = UCase$(Left$(pstrText, 1)) & LCase$(Mid$(pstrText, 2, Len(pstrText)))
End If
Capitalize = pstrText
End Function

Private Sub CheckLst()

  If lstmain.ListItems.Count = 0 Then
  mnuEditGenCode.Enabled = False
  mnuEditCopy.Enabled = False
  mnuEditRemove.Enabled = False
  mnuEditRemoveAll.Enabled = False
  cmdAddList.Enabled = True
  cmdRemoveSel.Enabled = False
  cmdRemoveAll.Enabled = False
  cmdGenerate.Enabled = False
  End If
  
  If Not lstmain.ListItems.Count = 0 Then
  cmdGenerate.Enabled = True
  mnuEditGenCode.Enabled = True
  mnuEditCopy.Enabled = True
  mnuEditRemove.Enabled = True
  mnuEditRemoveAll.Enabled = True
  cmdRemoveSel.Enabled = True
  cmdRemoveAll.Enabled = True
  End If
  
  If DocChanged = True Then
  mnuFileSave.Enabled = True
  End If
  
  If DocChanged = False Then
  mnuFileSave.Enabled = False
  End If
  
  If lstmain.ListItems.Count = 0 Then
  mnuFileSave.Enabled = False
  Else
  mnuFileSave.Enabled = True
  End If
         
End Sub

Private Sub ClearTxt()
txtProperty.Text = ""
cmbType.Text = ""
txtVariable.Text = ""
txtDefValue.Text = ""
rCode.Text = ""

End Sub

Private Function countComa()
   Dim i As Long
   Dim r As Long
   Dim LV As LV_ITEM
   
  'a string to build the msgbox text with
   Dim b As String

  'iterate through each item, checking its item state
   For i = 0 To lstmain.ListItems.Count
      r = SendMessage(lstmain.hwnd, LVM_GETITEMSTATE, i, ByVal LVIS_STATEIMAGEMASK)
     'when an item is checked, the LVM_GETITEMSTATE call
     'returns 8192 (&H2000&).
      If (r And &H2000&) Then
         'it is checked, so pad the LV_ITEM string members
         With LV
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
         End With
        'and retrieve the value (text) of the checked item
         Call SendMessage(lstmain.hwnd, LVM_GETITEMTEXT, i, LV)
         b = b & CStr(i) & ","
      End If
   Next
   countComa = b
End Function

Private Function FindLVCHKED()
Dim CharCount As String
Dim Char As String
Char = ","
    ' returns 5 but 6 if +1
    CharCount = Occurs(countComa, Char) '+ 1
If CharCount <= 0 Then
CharCount = 0
FindLVCHKED = "0"
Else
FindLVCHKED = CharCount
End If
End Function

Private Sub listcount()
If lstmain.ListItems.Count >= 1 Then
lbllstcount.Caption = FindLVCHKED & " of " & lstmain.ListItems.Count & " are selected"
Else
lbllstcount.Caption = "List is empty"
End If

End Sub

Private Sub LoadArray()
On Error GoTo exits:
Dim LstIcon As Integer
Dim the_array() As String
Dim list_item As ListItem
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim r As Long
Dim C As Long

Dim Cancel As Boolean
On Error GoTo errorhandler
Cancel = False

CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files|*.*"
CDL1.CancelError = True
CDL1.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
CDL1.ShowOpen

If Not Cancel Then
        file_name = CDL1.FileName
        docname = file_name
        Me.Caption = App.Title & " " & docname
        DocChanged = False
End If
GoTo loadl:
' -------------------
errorhandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If

loadl:
On Error GoTo exits:
    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)

    ' Copy the data into the array.
    For r = 0 To num_rows
        one_line = Split(lines(r), ",")
        For C = 0 To num_cols
            the_array(r, C) = one_line(C)
        Next C
    Next r
    
    ' Prove we have the data loaded.

For i = 1 To r
    If xx >= r Then xx = 0
Dim cb As String
    cb = the_array(xx, 2)

If cb = "Boolean" Then
LstIcon = 3
ElseIf cb = "Byte" Then
LstIcon = 3
ElseIf cb = "Currency" Then
LstIcon = 3
ElseIf cb = "Date" Then
LstIcon = 3
ElseIf cb = "Double" Then
LstIcon = 3
ElseIf cb = "Integer" Then
LstIcon = 3
ElseIf cb = "Long" Then
LstIcon = 3
ElseIf cb = "New" Then
LstIcon = 3
ElseIf cb = "OLE_CANCELBOOL" Then
LstIcon = 3
ElseIf cb = "OLE_COLOR" Then
LstIcon = 3
ElseIf cb = "OLE_HANDLE" Then
LstIcon = 3
ElseIf cb = "OLE_OPTEXCLUSIVE" Then
LstIcon = 3
ElseIf cb = "Single" Then
LstIcon = 3
ElseIf cb = "StdFont" Then
LstIcon = 4
ElseIf cb = "StdPicture" Then
LstIcon = 4
ElseIf cb = "String" Then
LstIcon = 3
ElseIf cb = "Variant" Then
LstIcon = 3
End If
    Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
    list_item.SmallIcon = LstIcon
    list_item.SubItems(1) = the_array(xx, 1)
    list_item.SubItems(2) = the_array(xx, 2)
    list_item.SubItems(3) = the_array(xx, 3)
    list_item.SubItems(4) = the_array(xx, 4)
xx = xx + 1
Next i

exits:
DocChanged = True
End Sub

Private Sub SaveNow()
On Error Resume Next
rCode.Text = ""
For i = 1 To lstmain.ListItems.Count
rCode.Text = rCode.Text & lstmain.ListItems.Item(i)
rCode.Text = rCode.Text & "," & lstmain.ListItems.Item(i).SubItems(1)
rCode.Text = rCode.Text & "," & lstmain.ListItems.Item(i).SubItems(2)
rCode.Text = rCode.Text & "," & lstmain.ListItems.Item(i).SubItems(3)
rCode.Text = rCode.Text & "," & lstmain.ListItems.Item(i).SubItems(4)
If i = lstmain.ListItems.Count Then GoTo save:
rCode.Text = rCode.Text & vbNewLine
Next

save:

End Sub

Private Sub cmbType_Click()
cmbType.SelLength = 0
Select Case UCase$(cmbType.Text)
    Case "INTEGER", "LONG", "SINGLE", "DOUBLE"
        txtDefValue.Text = "0"
    Case "BOOLEAN"
        txtDefValue.Text = "False"
        
End Select

End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
ComboKeyPress cmbType, KeyAscii
End Sub

Private Sub cmdAddList_Click()
  If cmbType.Text = "" Or cmbType.Text = vbNullString Then
    MsgBox "Debe indicar Tipo de Dato", vbOKOnly
    cmbType.SetFocus
    Exit Sub
  End If
  
  Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
  list_item.SmallIcon = SetListIcon(cmbType)
  list_item.SubItems(1) = txtProperty 'txtmain(0)
  list_item.SubItems(2) = cmbType
  list_item.SubItems(3) = txtVariable 'txtmain(2)
  list_item.SubItems(4) = txtDefValue 'txtmain(3)
  
  DocChanged = True
  listcount
  CheckLst
  txtProperty = Empty
  cmbType = Empty
  txtVariable = Empty
  txtDefValue = Empty
  rCode.Text = Empty
End Sub

Private Sub cmdGenerate_Click()
Dim sDIM As String, sSET As String
If FindLVCHKED <= 0 Then Exit Sub
    Dim xitem As Integer
    
    rCode.Text = Empty
    sSET = ""
    
    If chkReduced.Value = 1 Then
        rCode.Text = "'Variables-------------------" & vbNewLine
    Else
        rCode.Text = "Option Explicit" & vbNewLine & vbNewLine
    End If
    
    'Generate Dim/Private Declaractions
    If OptDIM(0).Value = True Then
        sDIM = "Dim "
    Else
        sDIM = "Private "
    End If
    For i = 0 To FindLVCHKED - 1
      xitem = Get_After_Comma(i, countComa)
      With lstmain.ListItems.Item(xitem + 1)
          rCode.Text = rCode.Text & sDIM & .SubItems(3) & " As " & .SubItems(2) & vbNewLine
      End With
    Next
    
    rCode.Text = rCode.Text & vbNewLine
    
    'Generate Get, Let, Set properties
    If chkReduced.Value = 1 Then
        rCode.Text = rCode.Text & "'Properties-------------------" & vbNewLine
    End If
    For i = 0 To FindLVCHKED - 1
      xitem = Get_After_Comma(i, countComa)
      With lstmain.ListItems.Item(xitem + 1)
          rCode.Text = rCode.Text & Generate(.SubItems(1), .SubItems(2), .SubItems(3)) & vbNewLine & vbNewLine
      End With
    Next
    
    'Generate UserControl_ReadProperties
    If chkReduced.Value = 1 Then
        rCode.Text = rCode.Text & "'ReadProperties Bag-------------------" & vbNewLine
    Else
        rCode.Text = rCode.Text & vbNewLine & "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine
        rCode.Text = rCode.Text & "   With PropBag" & vbNewLine
    End If
    
    For i = 0 To FindLVCHKED - 1
    xitem = Get_After_Comma(i, countComa)
        With lstmain.ListItems.Item(xitem + 1)
            If .SubItems(2) = "String" Then .SubItems(4) = """" & .SubItems(4) & """"
            If .SubItems(2) = "StdPicture" Or .SubItems(2) = "StdFont" Then sSET = "Set "
            rCode.Text = rCode.Text & vbTab & sSET & .SubItems(3) & " = .ReadProperty(" & """" & .SubItems(1) & """" & ", " & .SubItems(4) & ")" & vbNewLine
        End With
    Next
        
    If chkReduced.Value = 1 Then
        rCode.Text = rCode.Text & vbNewLine & "'WriteProperties Bag-------------------" & vbNewLine
    Else
        rCode.Text = rCode.Text & "   End With" & vbNewLine & "End Sub" & vbNewLine & vbNewLine
        'Generate UserControl_WriteProperties
        rCode.Text = rCode.Text & "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)" & vbNewLine
        rCode.Text = rCode.Text & "   With PropBag" & vbNewLine
    End If
    
    For i = 0 To FindLVCHKED - 1
    xitem = Get_After_Comma(i, countComa)
        With lstmain.ListItems.Item(xitem + 1)
            If .SubItems(2) = "String" Then .SubItems(4) = """" & .SubItems(4) & """"
            rCode.Text = rCode.Text & vbTab & "Call .WriteProperty(" & """" & .SubItems(1) & """" & ", " & .SubItems(3) & ", " & .SubItems(4) & ")" & vbNewLine
            '.SubItems(4) = Replace(.SubItems(4), """", "")
        End With
    Next
    
    If chkReduced.Value = 0 Then
        rCode.Text = rCode.Text & "   End With" & vbNewLine
        rCode.Text = rCode.Text & "End Sub"
    End If
    
  With rCode
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = &HF2F2F2
  End With

End Sub

Private Sub cmdInvertSel_Click()
EnhListView_InvertAllChecks lstmain
listcount
End Sub

Private Sub cmdRemoveAll_Click()
    If MsgBox("Are you sure you want to delete all entries in the List?", _
    vbCritical + vbYesNo, App.Title) = vbNo _
    Then Exit Sub
    
    lstmain.ListItems.Clear
    listcount
    CheckLst
End Sub

Private Sub cmdRemoveSel_Click()
Dim i As Integer
Dim srtx As String
On Error Resume Next
Do
  For i = 0 To FindLVCHKED - 1
    srtx = Get_After_Comma(i, countComa)
    lstmain.ListItems.Remove lstmain.ListItems.Item(srtx + 1).index
    listcount
    CheckLst
    rCode.Text = Empty
  Next i
Loop Until FindLVCHKED <= 0
End Sub

Private Sub cmdSelAll_Click()
EnhLitView_CheckAllItems lstmain
listcount
CheckLst
End Sub

Private Sub cmdUnSel_Click()
EnhLitView_UnCheckAllItems lstmain
listcount
CheckLst
End Sub

Private Sub Form_Load()
Dim list_item As ListItem

    Set X.list = lstmain
    X.addcolumn "ID", "id", 700, True, False
    X.addcolumn "Property Name", "pname", 1640, True, False
    X.addcolumn "Type Var", "tvar", 1600, False, True
    X.addcolumn "Container Var", "cvar", 1600, False, True
    X.addcolumn "Default Value", "defvalue", 1600, False, False
    lstmain.SmallIcons = imgSmall
    
    If cmbType.listcount = 0 Then LoadCombo cmbType
    
    ShowHeaderIcon 0, 0, True
    
    tHt.iItem = -1
    ' set lvVSS to set nodes for project.
    Call ListView_FullRowSelect(lstmain)
    Call ListView_GridLines(lstmain)
   
   lstmain.Refresh
   CheckLst
   listcount
   
    Call SendMessage(lstmain.hwnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_CHECKBOXES, ByVal True)

   Set TT = New CTooltip
   TT.Style = TTBalloon
   TT.Icon = TTIconInfo
    lstmain.Refresh
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
rCode.Move 75, 4080, (Me.ScaleWidth - 150), (Me.ScaleHeight - (rCode.Top + 75))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DocChanged = True Then
    
    Select Case MsgBox( _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, "ActiveX Coder 4")
    
    Case vbYes
        mnuFileS_Click
    Case vbNo
        Unload frmMain
    Case vbCancel
        Cancel = True
    
    End Select

End If
End Sub

Private Sub lstmain_Click()
listcount
End Sub

Private Sub lstmain_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

   Dim i As Long
   Static sOrder
   
   sOrder = Not sOrder
   
  'Use default sorting to sort the items in the list
   lstmain.SortKey = ColumnHeader.index - 1
   lstmain.SortOrder = Abs(sOrder)
   lstmain.Sorted = True
   
  'clear the image from the headers not
  'currently selected, and update the
  'header clicked
   For i = 0 To 4
      
     'if this is the index of the header clicked
      If i = lstmain.SortKey Then
      
           'ShowHeaderIcon colNo, imgIndex, showFlag
            ShowHeaderIcon lstmain.SortKey, _
                           lstmain.SortOrder, _
                           True
                           
      Else: ShowHeaderIcon i, 0, False
      End If
   
   Next
   
End Sub

Private Sub lstmain_DblClick()
Call mnuLstMainEdit_Click
End Sub

Private Sub lstmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
    tHt = ListView_HitTest(lstmain, X, Y)
        
    If Button <> 2 Then Exit Sub
    
    If tHt.iItem = -1 Then
        mnuLstMainRemoveAll.Enabled = False
        mnuLstMainEdit.Enabled = False
        mnuLstMainRemove.Enabled = False
    Else
        mnuLstMainRemove.Enabled = True
        mnuLstMainEdit.Enabled = True
        mnuLstMainRemoveAll.Enabled = True
        lstmain.ListItems(tHt.iItem + 1).Selected = True
    End If
    
    PopupMenu mnuLstMain
End Sub

Private Sub lstmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   Dim lvs As ListItem
   lvhti.pt.X = X / Screen.TwipsPerPixelX
   lvhti.pt.Y = Y / Screen.TwipsPerPixelY
   lItemIndex = SendMessage(lstmain.hwnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
      Set lvs = lstmain.ListItems(m_lCurItemIndex)
         TT.Title = "Property Info "
         TT.TipText = lstmain.ColumnHeaders.Item(2) & ": " & lvs.SubItems(1) _
         & vbCrLf & lstmain.ColumnHeaders.Item(3) & ": " & lvs.SubItems(2) _
         & vbCrLf & lstmain.ColumnHeaders.Item(4) & ": " & lvs.SubItems(3) _
         & vbCrLf & lstmain.ColumnHeaders.Item(5) & ": " & lvs.SubItems(4)
         TT.Create lstmain.hwnd
      End If
   End If
End Sub

Private Sub lstmain_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo Err
Dim LstIcon As Integer
Dim the_array() As String
Dim list_item As ListItem
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim r As Long
Dim C As Long

    ' Load the file.
    file_name = Data.Files(1)
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)

    ' Copy the data into the array.
    For r = 0 To num_rows
        one_line = Split(lines(r), ",")
        For C = 0 To num_cols
            the_array(r, C) = one_line(C)
        Next C
    Next r
    
    ' Prove we have the data loaded.
For i = 1 To r
        If xx >= r Then xx = 0
Dim cb As String
    cb = the_array(xx, 2)

If cb = "Boolean" Then
LstIcon = 3
ElseIf cb = "Byte" Then
LstIcon = 3
ElseIf cb = "Currency" Then
LstIcon = 3
ElseIf cb = "Date" Then
LstIcon = 3
ElseIf cb = "Double" Then
LstIcon = 3
ElseIf cb = "Integer" Then
LstIcon = 3
ElseIf cb = "Long" Then
LstIcon = 3
ElseIf cb = "New" Then
LstIcon = 3
ElseIf cb = "OLE_CANCELBOOL" Then
LstIcon = 3
ElseIf cb = "OLE_COLOR" Then
LstIcon = 3
ElseIf cb = "OLE_HANDLE" Then
LstIcon = 3
ElseIf cb = "OLE_OPTEXCLUSIVE" Then
LstIcon = 3
ElseIf cb = "Single" Then
LstIcon = 3
ElseIf cb = "StdFont" Then
LstIcon = 4
ElseIf cb = "StdPicture" Then
LstIcon = 4
ElseIf cb = "String" Then
LstIcon = 3
ElseIf cb = "Variant" Then
LstIcon = 3
End If
    Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
    list_item.SmallIcon = LstIcon
    list_item.SubItems(1) = the_array(xx, 1)
    list_item.SubItems(2) = the_array(xx, 2)
    list_item.SubItems(3) = the_array(xx, 3)
    list_item.SubItems(4) = the_array(xx, 4)
xx = xx + 1
Next i

listcount
CheckLst
DocChanged = True
    Exit Sub
Err:
    MsgBox "The File could not be loaded", vbExclamation
End Sub

Private Sub mnuEditAdd_Click()
cmdAddList_Click
End Sub

Private Sub mnuEditCopy_Click()
  Clipboard.Clear
  Clipboard.SetText rCode.Text
End Sub

Private Sub mnuEditGen_Click()
cmdGenerate_Click
End Sub

Private Sub mnuEditRemove_Click()
cmdRemoveSel_Click
End Sub

Private Sub mnuEditRemoveAll_Click()
cmdRemoveAll_Click
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileImp_Click()
LoadArray
listcount
CheckLst
End Sub

Private Sub mnuFileLoad_Click()
mnuFileNew_Click
LoadArray
listcount
CheckLst
End Sub

Private Sub mnuFileNew_Click()
Dim Cancel As Integer

If DocChanged = False Then
    DocChanged = False
    lstmain.ListItems.Clear
    rCode.Text = ""
    listcount
    CheckLst
    ClearTxt
Else
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, "ActiveX Coder 3")
    
    Case vbYes
        mnuFileS_Click
    Case vbNo
        DocChanged = False
        lstmain.ListItems.Clear
        rCode.Text = ""
        listcount
        CheckLst
        ClearTxt
    Case vbCancel
        Cancel = True
    
    End Select
End If

End Sub


Private Sub mnuFileS_Click()
Call SaveNow
If docname = "" Then
    mnuFileSS_Click
Else
rCode.SaveFile docname, rtfText
DocChanged = False
End If

End Sub

Private Sub mnuFileSS_Click()
Call SaveNow
Dim Cancel As Boolean
On Error GoTo errorhandler
Cancel = False

CDL1.DefaultExt = ".txt"
CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
CDL1.CancelError = True
CDL1.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

CDL1.ShowSave

If Not Cancel Then
    If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
        rCode.SaveFile CDL1.FileName, rtfRTF
    Else
        rCode.SaveFile CDL1.FileName, rtfText
    End If
    rCode.FileName = CDL1.FileName
    docname = CDL1.FileName
    Me.Caption = App.Title & " " & docname
    DocChanged = False
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If

End Sub

Private Sub mnuLstMainEdit_Click()
On Error Resume Next
frmRow.open_dlg lstmain.SelectedItem.SubItems(1), _
                lstmain.SelectedItem.SubItems(2), _
                lstmain.SelectedItem.SubItems(3), _
                lstmain.SelectedItem.SubItems(4)
End Sub

Private Sub mnuLstMainRemove_Click()
cmdRemoveSel_Click
End Sub

Private Sub mnuLstMainRemoveAll_Click()
cmdRemoveAll_Click
End Sub

Private Sub txtProperty_Change()
    txtVariable.Text = "m_" & Trim$(txtProperty.Text)
    txtDefValue.Text = Trim$(txtProperty.Text)
End Sub

Private Sub txtProperty_LostFocus()
  txtProperty.Text = Capitalize(txtProperty.Text)
End Sub


