VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'****************************************************************************************
'* clsClassList - Editable cells for ListView Control                                   *
'****************************************************************************************
'*                                                                                      *
'* Created By:      Aaron Thorp                                                         *
'* Contact:         in_my_datto@hotmail.com                                             *
'* Last Modified:   4th January 2004                                                    *
'*                                                                                      *
'****************************************************************************************
'*                                                                                      *
'* This class will provide click-to-edit cells in listviews.                            *
'* To implement this class, the main form will need to have a listview and a textbox    *
'* and declare these to the class by settings the properties "list" and "textbox"       *
'*                                                                                      *
'****************************************************************************************
'* ISSUES/KNOWN BUGS:                                                                   *
'*                                                                                      *
'*  - There can only be one field that has GROW enabled.                                *
'*  - This class will not work for ListViews that have columns that need to be accessed *
'*    with Horizontal scrollbars.                                                       *
'*  - If you click off the listview after editing, it will not save the changes and     *
'*    clear the textbox until you click inside the listview again.                      *
'****************************************************************************************

Private WithEvents mvarlist As ListView
Attribute mvarlist.VB_VarHelpID = -1
Private WithEvents mvartext As TextBox
Attribute mvartext.VB_VarHelpID = -1

Private edit_item As ListItem

Private active_column As Integer
Private edit_active As Boolean

Private old_cell_value As String

Private Px As Single, Py As Single

Private Const SET_LabelEdit = lvwManual
Private Const SET_HideSelection = False
Private Const SET_FullRowSelect = True
Private Const SET_GridLines = True
Private Const SET_View = lvwReport

Private Column As clsColumnSettings
Private Columns As New Collection

Public Sub additem(item_text As String, ParamArray item_subitems() As Variant)

    Dim Item As ListItem, i As Integer
    
    Set Item = mvarlist.ListItems.Add(, , item_text)
    
    For i = 1 To UBound(item_subitems) + 1
        If i < mvarlist.ColumnHeaders.Count Then Item.SubItems(i) = item_subitems(i - 1)
    Next

End Sub

Public Sub addcolumn(col_name As String, col_key, col_width As Double, col_grow As Boolean, Optional col_edit As Boolean = False, Optional col_align As Integer = lvwColumnLeft, Optional col_show As Boolean = True)
'adds and defines the column header settings storing them in the clsColumnSettings collection

    Dim Col As ColumnHeader
    
    Set Column = New clsColumnSettings
           
    Column.colwidth = col_width
    Column.align = col_align
    Column.edit = col_edit
    Column.grow = col_grow
    Column.show = col_show
    
    Set Col = mvarlist.ColumnHeaders.Add(, col_key, col_name, col_width, col_align)
    
    Column.index = Col.index
    
    Columns.Add Column, col_key
    
End Sub

Public Property Set list(ByVal vData As Object)
    
    'define which listview control the class is controlling
    Set mvarlist = vData

End Property

Public Property Get list() As Object
    
    Set list = mvarlist

End Property


Private Sub Class_Initialize()

    edit_active = False
    Set Columns = New Collection

End Sub
