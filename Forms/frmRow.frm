VERSION 5.00
Begin VB.Form frmRow 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Row Modifier"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2280
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function open_dlg(par1, par2, par3, par4)
  Text3(0) = par1
  Combo1 = par2
  Text3(2) = par3
  Text3(3) = par4
  Me.show vbModal, frmMain
End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
ComboKeyPress Combo1, KeyAscii
End Sub

Private Sub Combo1_LostFocus()
Combo1.SelLength = 0
End Sub

Private Sub Command1_Click()
On Error GoTo funExit
frmMain.lstmain.SelectedItem.SmallIcon = SetListIcon(Me.Combo1)
ListView_ModifyRow frmMain.lstmain, Text3(0), frmMain.lstmain.SelectedItem, , , 1
ListView_ModifyRow frmMain.lstmain, Combo1, , frmMain.lstmain.SelectedItem.index - 1, , 2
ListView_ModifyRow frmMain.lstmain, Text3(2), , frmMain.lstmain.SelectedItem.index - 1, , 3
ListView_ModifyRow frmMain.lstmain, Text3(3), , frmMain.lstmain.SelectedItem.index - 1, , 4

funExit:
Me.Hide
End Sub

Private Sub Form_Load()
LoadCombo Combo1
End Sub
