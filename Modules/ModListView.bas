Attribute VB_Name = "modListView"
Option Explicit

Public Const LVM_FIRST = &H1000
Public Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2

Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_IMAGE = &H800
Public Const HDF_STRING = &H4000

Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Public Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

' styles for listview
Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H37
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + &H36
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)

Public Const MAX_PATH = 260
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVIF_STATE = &H8
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVIS_STATEIMAGEMASK As Long = &HF000

' hittest constants
Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

' edit subitem constants
Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
Public Const LVIR_SELECTBOUNDS = 3

Public Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   iImage As Long
   iOrder As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type LVHITTESTINFO
    pt As POINTAPI
    lFlags As Long
    lItem As Long
    lSubItem As Long
End Type

Public Type LV_ITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   state        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
lpRect As RECT) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
   (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, _
   ByVal lpszClass As String, ByVal lpszWindow As String) As Long

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   

Public Sub ShowHeaderIcon(colNo As Long, imgIconNo As Long, showImage As Long)

   Dim hHeader As Long
   Dim HD As HD_ITEM
   
  'get a handle to the listview header component
   hHeader = SendMessage(frmMain.lstmain.hwnd, LVM_GETHEADER, 0, ByVal 0)
   
  'set up the required structure members
   With HD
      .mask = HDI_IMAGE Or HDI_FORMAT
      .pszText = frmMain.lstmain.ColumnHeaders(colNo + 1).Text
      
       If showImage Then
         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
         .iImage = imgIconNo
       Else
         .fmt = HDF_STRING
      End If

   End With
      
  'modify the header
   Call SendMessage(hHeader, HDM_SETITEM, colNo, HD)
End Sub

Public Function Generate(name As String, typ As String, variable As String) As String
If typ = "StdPicture" Or typ = "StdFont" Then
    Generate = "Public Property " & "Get " & name & "() as " & typ & vbNewLine & _
                vbTab & "Set " & name & " = " & variable & vbNewLine & _
                "End Property" & vbNewLine
                
    If typ = "StdPicture" Then Generate = Generate & vbNewLine & _
               "Public Property " & "Let " & name & "(ByVal New" & name & " as " & typ & ")" & vbNewLine & _
               vbTab & "Set " & variable & " = " & "New" & name & " " & vbNewLine & _
               vbTab & "PropertyChanged " & """" & name & """" & vbNewLine & _
               vbTab & IIf(frmMain.chkRefresh.Value = 1, "Refresh", "UserControl_Paint") & vbNewLine & _
               "End Property" & vbNewLine
               
    Generate = Generate & vbNewLine & "Public Property " & "Set " & name & "(ByVal New" & name & " as " & typ & ")" & vbNewLine & _
               vbTab & "Set " & variable & " = " & "New" & name & " " & vbNewLine & _
               vbTab & "PropertyChanged " & """" & name & """" & vbNewLine & _
               "End Property"
Else
    Generate = "Public Property " & "Get " & name & "() as " & typ & vbNewLine & _
               vbTab & name & " = " & variable & vbNewLine & _
               "End Property" & vbNewLine
               
    Generate = Generate & vbNewLine & _
               "Public Property " & "Let " & name & "(ByVal New" & name & " as " & typ & ")" & vbNewLine & _
               vbTab & variable & " = " & "New" & name & " " & vbNewLine & _
               vbTab & "PropertyChanged " & """" & name & """" & vbNewLine & _
               IIf(frmMain.chkRefresh.Value = 1, vbTab & "Refresh" & vbNewLine, vbNewLine) & _
               "End Property"

End If
End Function

Public Function ListView_ModifyRow(ByRef lstControl As ComctlLib.ListView, _
                                   Optional ByVal strValue As String, _
                                   Optional ByRef lstRowToModify As ComctlLib.ListItem, _
                                   Optional ByVal intRowIndex As Integer = -1, _
                                   Optional ByVal strRowKey As String, _
                                   Optional ByVal intColumnIndex As Integer = -1, _
                                   Optional ByRef Return_ErrNum As Long, _
                                   Optional ByRef Return_ErrDesc As String) As Boolean
On Error Resume Next
  
  ' Clear return variables
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  strRowKey = Trim(strRowKey)
  If lstControl Is Nothing Then GoTo InvalidParameter
  If intColumnIndex <> -1 Then If intColumnIndex > lstControl.ColumnHeaders.Count Or intColumnIndex < 0 Then GoTo InvalidParameter
  If intRowIndex <> -1 Then
    intRowIndex = intRowIndex + 1
    If intRowIndex > lstControl.ListItems.Count Or intRowIndex < 1 Then GoTo InvalidParameter
  End If
  If lstRowToModify Is Nothing And intRowIndex < 0 And strRowKey = "" Then GoTo InvalidParameter
  
  ' If the user wants to edit a column in the row, then do so
  If intColumnIndex > 0 Then
    If Not lstRowToModify Is Nothing Then
      If intColumnIndex = 0 Then lstRowToModify.Text = strValue Else lstRowToModify.SubItems(intColumnIndex) = strValue
    ElseIf strRowKey <> "" Then
      If intColumnIndex = 0 Then lstControl.ListItems.Item(strRowKey).Text = strValue Else lstControl.ListItems.Item(strRowKey).SubItems(intColumnIndex) = strValue
    ElseIf intRowIndex > 0 Then
      If intColumnIndex = 0 Then lstControl.ListItems.Item(intRowIndex).Text = strValue Else lstControl.ListItems.Item(intRowIndex).SubItems(intColumnIndex) = strValue
    End If
    
  ' The user wants to edit the row's valud (first column)
  Else
    If Not lstRowToModify Is Nothing Then
      lstRowToModify.Text = strValue
    ElseIf strRowKey <> "" Then
      lstControl.ListItems.Item(strRowKey).Text = strValue
    ElseIf intRowIndex > 0 Then
      lstControl.ListItems.Item(intRowIndex).Text = strValue
    End If
  End If
  
  ' Check for errors
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  If Return_ErrNum = 0 Then ListView_ModifyRow = True
  
  Exit Function
  
InvalidParameter:
  
  Return_ErrNum = -1
  Return_ErrDesc = "Invalid parameter(s) passed to the 'ListView_ModifyRow' function"
  
End Function

Public Sub ListView_FullRowSelect(ByRef ListView As ListView)

    Dim lStyle As Long
    lStyle = SendMessage(ListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or LVS_EX_FULLROWSELECT Or lStyle
    Call SendMessage(ListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub

Public Sub ListView_GridLines(ByRef ListView As ListView)
    
    Dim lStyle As Long
    lStyle = SendMessage(ListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Or LVS_EX_GRIDLINES
    Call SendMessage(ListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)

End Sub

Public Function ListView_HitTest(ListView As ListView, X As Single, Y As Single) As LVHITTESTINFO
    
    Dim lRet As Long
    Dim lX As Long
    Dim lY As Long

    'x and y are in twips; convert them to pixels for the API call
    lX = X / Screen.TwipsPerPixelX
    lY = Y / Screen.TwipsPerPixelY

    Dim tHitTest As LVHITTESTINFO
    
    With tHitTest
        .lFlags = 0
        .lItem = 0
        .lSubItem = 0
        .pt.X = lX
        .pt.Y = lY
    End With
    
    'return the filled Structure to the routine
    lRet = SendMessage(ListView.hwnd, LVM_SUBITEMHITTEST, 0, tHitTest)
    
    ListView_HitTest = tHitTest
    
End Function

Public Sub ComboKeyPress(Combo1 As ComboBox, KeyAscii As Integer)
    Dim strSearchText As String
    Dim strEnteredText As String
    Dim intLength As Integer
    Dim intIndex As Integer
    Dim intCounter As Integer
    
On Error GoTo errorhandler

With Combo1

    If .SelStart > 0 Then
        strEnteredText = Left$(.Text, .SelStart)
    End If

    Select Case KeyAscii
        Case vbKeyReturn
        If .ListIndex > -1 Then
            .SelStart = 0
            .SelLength = Len(.list(.ListIndex))
            Exit Sub
        End If
        
        Case vbKeyEscape, vbKeyDelete
        .Text = ""
        KeyAscii = 0
        Exit Sub
        
        Case vbKeyBack
        If Len(strEnteredText) > 1 Then
            strSearchText = LCase(Left(strEnteredText, Len(strEnteredText) - 1))
        Else
            strEnteredText = ""
            KeyAscii = 0
            .Text = ""
            Exit Sub
        End If
        
        Case Else
        strSearchText = LCase(strEnteredText & Chr(KeyAscii))
        
    End Select
    
    intIndex = -1
    intLength = Len(strSearchText)
    
    For intCounter = 0 To .listcount - 1
      If LCase(Left(.list(intCounter), intLength)) = strSearchText Then
          intIndex = intCounter
          Exit For
      End If
    Next intCounter
    
    If intIndex > -1 Then
        .ListIndex = intIndex
        .SelStart = Len(strSearchText)
        .SelLength = Len(.list(intIndex)) - Len(strSearchText)
        KeyAscii = 0
    Else
        'Beep
    End If
End With

'KeyAscii = 0
Exit Sub
errorhandler:
'KeyAscii = 0
Beep
End Sub

Public Sub LoadCombo(Combo1 As ComboBox)
    With Combo1
        .additem "Boolean"
        .additem "Byte"
        .additem "Collection"
        .additem "Currency"
        .additem "Date"
        .additem "Double"
        .additem "Integer"
        .additem "Long"
        .additem "New"
        .additem "OLE_CANCELBOOL"
        .additem "OLE_COLOR"
        .additem "OLE_HANDLE"
        .additem "OLE_OPTEXCLUSIVE"
        .additem "Single"
        .additem "StdFont"
        .additem "StdPicture"
        .additem "String"
        .additem "Variant"
        .Text = ""
    End With
End Sub

Function Get_After_Comma(ByVal intCommaNumber As Integer, ByVal strString As String) As String
    Dim intIndex As Integer
    Dim intStartOfString As Integer
    Dim intEndOfString As Integer
    Dim boolNotFound As Integer
    'check for intCommaNumber = 0--i.e. firs
    '     t one

    If (intCommaNumber = 0) Then
        Get_After_Comma = Left$(strString, InStr(strString, ",") - 1)
    Else
        'not the first one init start of string on first comma
        intStartOfString = InStr(strString, ",")
        'place start of string after intCommaNum ber-th comma (-1 since already did one
        boolNotFound = 0

        For intIndex = 1 To intCommaNumber - 1
            'get next comma
            intStartOfString = InStr(intStartOfString + 1, strString, ",")
            'check for not found

            If (intStartOfString = 0) Then
                boolNotFound = 1
            End If
        Next intIndex
        'put start of string past 1st comma
        intStartOfString = intStartOfString + 1
        'check for ending in a comma

        If (intStartOfString > Len(strString)) Then
            boolNotFound = 1
        End If

        If (boolNotFound = 1) Then
            Get_After_Comma = "NOT FOUND"
        Else
            intEndOfString = InStr(intStartOfString, strString, ",")
            'check for no second comma (i.e. end of string)

            If (intEndOfString = 0) Then
                intEndOfString = Len(strString) + 1
            Else
                intEndOfString = intEndOfString - 1
            End If
            Get_After_Comma = Mid$(strString, intStartOfString, intEndOfString - intStartOfString + 1)
        End If
    End If
End Function


Public Function Occurs(ByVal strtochk As String, ByVal searchstr As String) As Long
    ' remember SPLIT returns a zero-based array
    Occurs = UBound(Split(strtochk, searchstr)) '+1
End Function

'Set Itm with var type icons

Public Function SetListIcon(ComboBox As ComboBox)
Dim cb As ComboBox
Set cb = ComboBox

Select Case cb
    Case "StdFont", "StdPicture", "Collection"
        SetListIcon = 4
    Case Else
        SetListIcon = 3
End Select

End Function

Public Function EnhLitView_CheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhLitView_CheckAllItems
    
    '________________________________________________________________________
    ' set function return to true
    EnhLitView_CheckAllItems = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LV_ITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '________________________________________________________________________
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessage(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhLitView_CheckAllItems:
    
    '________________________________________________________________________
    ' set function return to false
    EnhLitView_CheckAllItems = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhLitView_CheckAllItems"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_CheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function

Public Function EnhLitView_UnCheckAllItems( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhLitView_UnCheckAllItems
    
    '________________________________________________________________________
    ' set function return to true
    EnhLitView_UnCheckAllItems = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LV_ITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    '________________________________________________________________________
    lvState = IIf(False, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .state = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessage(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhLitView_UnCheckAllItems:
    
    '________________________________________________________________________
    ' set function return to false
    EnhLitView_UnCheckAllItems = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhLitView_UnCheckAllItems"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhLitView_UnCheckAllItems" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function

Public Function EnhListView_InvertAllChecks( _
                lstListViewName As ListView, _
                Optional bolShowErrors As Boolean) _
                As Boolean
    
    '________________________________________________________________________
    ' initiate error handler
    On Error GoTo err_EnhListView_InvertAllChecks
    
    '________________________________________________________________________
    ' set function return to true
    EnhListView_InvertAllChecks = True
    
    '________________________________________________________________________
    ' setup variables
    Dim LV          As LV_ITEM
    Dim r           As Long
    Dim lvCount     As Long
    Dim lvIndex     As Long
    
    '________________________________________________________________________
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        r = SendMessageLong(lstListViewName.hwnd, LVM_GETITEMSTATE, lvIndex, LVIS_STATEIMAGEMASK)
        With LV
            .mask = LVIF_STATE
            .stateMask = LVIS_STATEIMAGEMASK
            If r And &H2000& Then
                .state = &H1000
            Else
                .state = &H2000
            End If
        End With
        r = SendMessage(lstListViewName.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    '________________________________________________________________________
    ' exit before error handler
    Exit Function
    
'________________________________________________________________________
' deal with errors
err_EnhListView_InvertAllChecks:
    
    '________________________________________________________________________
    ' set function return to false
    EnhListView_InvertAllChecks = False
    '________________________________________________________________________
    ' if you want notification on an error
    If bolShowErrors = True Then
        MsgBox "Error" & Err.Number & vbTab & Err.Description, _
               vbOKOnly + vbInformation, _
               "Error in Function : EnhListView_InvertAllChecks"
    End If
    
    '________________________________________________________________________
    ' initiate debug
    Debug.Print Now & vbTab & "Error in function: EnhListView_InvertAllChecks" _
                & vbCrLf & _
                Err.Number & vbTab & Err.Description
    Debug.Assert False
    
    '________________________________________________________________________
    ' exit
    Exit Function
    
End Function


