VERSION 5.00
Begin VB.UserControl cAbout 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4260
   ScaleWidth      =   6285
   ToolboxBitmap   =   "cAbout.ctx":0000
   Begin VB.PictureBox PicBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   0
      ScaleHeight     =   3870
      ScaleWidth      =   5520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5550
      Begin VB.PictureBox PicCoded 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4380
         Picture         =   "cAbout.ctx":0532
         ScaleHeight     =   615
         ScaleWidth      =   1050
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3180
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.PictureBox picLogo2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4380
         Picture         =   "cAbout.ctx":2768
         ScaleHeight     =   240
         ScaleWidth      =   870
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   870
      End
      Begin VB.PictureBox PicLogo1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   1305
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblApp1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculadora++"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   1380
         TabIndex        =   7
         Top             =   1365
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "David Rojas Arraño"
         Height          =   195
         Left            =   2820
         TabIndex        =   6
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Axio.UK ®2019"
         Height          =   195
         Left            =   3090
         TabIndex        =   5
         Top             =   3375
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Axio Soft. && Tech. ©1997-2019"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   3570
         Width           =   2295
      End
      Begin VB.Label lblVers 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v0.0.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   3555
         TabIndex        =   3
         Top             =   1725
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1365
         X2              =   4060
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Label lblApp2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculadora++"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   420
         Left            =   1350
         TabIndex        =   8
         Top             =   1335
         Width           =   2265
      End
   End
End
Attribute VB_Name = "cAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eAlign
  [Left Align] = 0
  [Right Align] = 1
  [Center Align] = 2
End Enum

Public Enum PicPos
  [Top Pos] = 0
  [Down Pos] = 1
End Enum

Public Enum tBorder
  None = 0
  [Fixed Single] = 1
End Enum

Dim uWidth As Integer
Dim uHeight As Integer
Dim mAlign As eAlign
Dim vAlign As eAlign
Dim tAlign As eAlign
Dim iPicPos As PicPos
Dim mForeColor As OLE_COLOR
Dim vForeColor As OLE_COLOR
Dim tForeColor As OLE_COLOR
Dim tShadowColor As OLE_COLOR
Dim cBackColor As OLE_COLOR
Dim sBorder As tBorder
Dim tYPos As Integer
Dim m_TitleFont As Font

'Event Declarations:
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Property Variables:





Private Sub Create()
uWidth = UserControl.ScaleWidth
uHeight = UserControl.ScaleHeight
PicBack.Move 0, 0, uWidth, uHeight
picLogo2.Move uWidth - (picLogo2.Width + 100)
PicBack.BorderStyle = sBorder

Label1.Alignment = mAlign
Label2.Alignment = mAlign
Label3.Alignment = mAlign

Select Case mAlign
  Case Is = 0
    If PicCoded.Visible = True Then
      PicCoded.Move 100, Label1.Top
      Label3.Move PicCoded.Width + 200, uHeight - (Label3.Height + 100)
      Label2.Move PicCoded.Width + 200, uHeight - (Label3.Height + Label2.Height + 100)
      Label1.Move PicCoded.Width + 200, uHeight - (Label3.Height + Label2.Height + Label1.Height + 100)
    Else
      Label3.Move 100, uHeight - (Label3.Height + 100)
      Label2.Move 100, uHeight - (Label3.Height + Label2.Height + 100)
      Label1.Move 100, uHeight - (Label3.Height + Label2.Height + Label1.Height + 100)
    End If
  Case Is = 1
    If PicCoded.Visible = True Then
      PicCoded.Move uWidth - (PicCoded.Width + 100), Label1.Top
      Label3.Move uWidth - (Label3.Width + PicCoded.Width + 200), uHeight - (Label3.Height + 100)
      Label2.Move uWidth - (Label2.Width + PicCoded.Width + 200), uHeight - (Label3.Height + Label2.Height + 100)
      Label1.Move uWidth - (Label1.Width + PicCoded.Width + 200), uHeight - (Label3.Height + Label2.Height + Label1.Height + 100)
    Else
      Label3.Move uWidth - (Label3.Width + 100), uHeight - (Label3.Height + 100)
      Label2.Move uWidth - (Label2.Width + 100), uHeight - (Label3.Height + Label2.Height + 100)
      Label1.Move uWidth - (Label1.Width + 100), uHeight - (Label3.Height + Label2.Height + Label1.Height + 100)
    End If
  Case Is = 2
    PicCoded.Visible = False
    Label3.Move (uWidth / 2) - (Label3.Width / 2), uHeight - (Label3.Height + 100)
    Label2.Move (uWidth / 2) - (Label2.Width / 2), uHeight - (Label3.Height + Label2.Height + 100)
    Label1.Move (uWidth / 2) - (Label1.Width / 2), uHeight - (Label3.Height + Label2.Height + Label1.Height + 100)

End Select

Line1.X1 = 250
Line1.X2 = uWidth - 250
Line1.Y1 = uHeight / 2
Line1.Y2 = uHeight / 2

Select Case vAlign
  Case Is = 0
    lblVers.Move 250, Line1.Y2 + 50
    
  Case Is = 1
    lblVers.Move uWidth - (lblVers.Width + 250), Line1.Y2 + 50
    
  Case Is = 2
    lblVers.Move uWidth / 2 - (lblVers.Width / 2), Line1.Y2 + 50
    
End Select

Select Case tAlign
  Case Is = 0
      lblApp1.Move 250, tYPos + 30, uWidth - 250
      lblApp2.Move 280, tYPos, uWidth - 250
  
  Case Is = 1
      lblApp1.Move uWidth - (lblApp1.Width + 250), tYPos + 30, uWidth - 250
      lblApp2.Move uWidth - (lblApp2.Width + 280), tYPos, uWidth - 250
      
  Case Is = 2
      lblApp1.Move uWidth / 2 - (lblApp1.Width / 2), tYPos + 30, uWidth - 250
      lblApp2.Move uWidth / 2 - (lblApp2.Width / 2), tYPos, uWidth - 250
      
End Select

End Sub

Private Sub UserControl_Initialize()
lblVers.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision & "  "
lblApp1.Caption = App.ProductName & "  "
lblApp2.Caption = App.ProductName & "  "
lblApp1.AutoSize = True
lblApp2.AutoSize = True
Label1.Caption = App.LegalCopyright & "  "
Label2.Caption = App.LegalTrademarks & "  "
Label3.Caption = App.CompanyName & "  "

End Sub

Private Sub UserControl_InitProperties()
vAlign = [Right Align]
tAlign = [Left Align]
mAlign = [Right Align]
iPicPos = [Top Pos]
cBackColor = vbWhite
tForeColor = &HFF0000
tShadowColor = &HFFC0C0
vForeColor = vbBlack
mForeColor = vbBlack
sBorder = None
tYPos = UserControl.ScaleHeight / 2 - (lblApp1.Height)
Set m_TitleFont = Ambient.Font

End Sub

Private Sub UserControl_Resize()
tYPos = UserControl.ScaleHeight / 2 - (lblApp1.Height)
Call Create
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  PicBack.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  PicBack.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  mAlign = PropBag.ReadProperty("AlignMarks", 1)
  tAlign = PropBag.ReadProperty("AlignTitle", 0)
  vAlign = PropBag.ReadProperty("AlignVersion", 1)
  tForeColor = PropBag.ReadProperty("TitleForecolor", &HFF0000)
  tShadowColor = PropBag.ReadProperty("TitleShadowColor", &HFFC0C0)
  mForeColor = PropBag.ReadProperty("MarksForeColor", &H80000012)
  vForeColor = PropBag.ReadProperty("VersionForeColor", &H808080)
  Set Label1.Font = PropBag.ReadProperty("MarksFont", Ambient.Font)
  Set lblVers.Font = PropBag.ReadProperty("VersionFont", Ambient.Font)
  
  Set lblApp1.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
  Set lblApp2.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
  Set m_TitleFont = PropBag.ReadProperty("TitleFont", Ambient.Font)

  Set Picture = PropBag.ReadProperty("BackPicture", Nothing)
  Set Picture = PropBag.ReadProperty("LogoPicture", Nothing)
  PicCoded.Enabled = PropBag.ReadProperty("VisibleCoded", False)
  tYPos = PropBag.ReadProperty("TitleYPos", UserControl.Height / 2 - (lblApp1.Height))
End Sub

Private Sub UserControl_Show()
Call Create
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", PicBack.BackColor, &H80000005)
  Call PropBag.WriteProperty("BorderStyle", PicBack.BorderStyle, 1)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("AlignMarks", mAlign, 1)
  Call PropBag.WriteProperty("AlignTitle", tAlign, 0)
  Call PropBag.WriteProperty("AlignVersion", vAlign, 1)
  Call PropBag.WriteProperty("TitleForecolor", tForeColor, &HFF0000)
  Call PropBag.WriteProperty("TitleShadowColor", tShadowColor, &HFFC0C0)
  Call PropBag.WriteProperty("MarksForeColor", mForeColor, &H80000012)
  Call PropBag.WriteProperty("VersionForeColor", vForeColor, &H808080)
  Call PropBag.WriteProperty("MarksFont", Label1.Font, Ambient.Font)
  Call PropBag.WriteProperty("VersionFont", lblVers.Font, Ambient.Font)
  'Call PropBag.WriteProperty("TitleFont", lblApp1.Font, Ambient.Font)
  Call PropBag.WriteProperty("TitleFont", m_TitleFont, Ambient.Font)
  Call PropBag.WriteProperty("BackPicture", Picture, Nothing)
  Call PropBag.WriteProperty("LogoPicture", Picture, Nothing)
  Call PropBag.WriteProperty("VisibleCoded", PicCoded.Enabled, False)
  Call PropBag.WriteProperty("TitleYPos", tYPos, UserControl.Height / 2 - (lblApp1.Height))
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get AlignMarks() As eAlign
Attribute AlignMarks.VB_Description = "Devuelve o establece la alineación de un control CheckBox u OptionButton, o el texto de un control."
  AlignMarks = mAlign
End Property

Public Property Let AlignMarks(ByVal New_AlignMarks As eAlign)
  mAlign = New_AlignMarks
  PropertyChanged "AlignMarks"
  Create
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblApp1,lblApp1,-1,Alignment
Public Property Get AlignTitle() As eAlign
Attribute AlignTitle.VB_Description = "Devuelve o establece la alineación de un control CheckBox u OptionButton, o el texto de un control."
  AlignTitle = tAlign
End Property

Public Property Let AlignTitle(ByVal New_AlignTitle As eAlign)
  tAlign = New_AlignTitle
  PropertyChanged "AlignTitle"
  Create
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblVers,lblVers,-1,Alignment
Public Property Get AlignVersion() As eAlign
Attribute AlignVersion.VB_Description = "Devuelve o establece la alineación de un control CheckBox u OptionButton, o el texto de un control."
  AlignVersion = vAlign
End Property

Public Property Let AlignVersion(ByVal New_AlignVersion As eAlign)
  vAlign = New_AlignVersion
  PropertyChanged "AlignVersion"
  Create
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=PicBack,PicBack,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  BackColor = PicBack.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  PicBack.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=PicBack,PicBack,-1,Picture
Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
  Set BackPicture = PicBack.Picture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
  Set PicBack.Picture = New_BackPicture
  PropertyChanged "BackPicture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=PicBack,PicBack,-1,BorderStyle
Public Property Get BorderStyle() As tBorder
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
  BorderStyle = sBorder
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As tBorder)
  sBorder = New_BorderStyle
  PropertyChanged "BorderStyle"
  Create
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=picLogo2,picLogo2,-1,Picture
Public Property Get LogoPicture() As Picture
Attribute LogoPicture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
  Set LogoPicture = PicLogo1.Picture
End Property

Public Property Set LogoPicture(ByVal New_LogoPicture As Picture)
  Set PicLogo1.Picture = New_LogoPicture
  PropertyChanged "LogoPicture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,Font
Public Property Get MarksFont() As Font
Attribute MarksFont.VB_Description = "Devuelve un objeto Font."
  Set MarksFont = Label1.Font
End Property

Public Property Set MarksFont(ByVal New_MarksFont As Font)
  Set Label1.Font = New_MarksFont
  Set Label2.Font = New_MarksFont
  Set Label3.Font = New_MarksFont
  PropertyChanged "MarksFont"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get MarksForeColor() As OLE_COLOR
Attribute MarksForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  MarksForeColor = Label1.ForeColor
End Property

Public Property Let MarksForeColor(ByVal New_MarksForeColor As OLE_COLOR)
  Label1.ForeColor() = New_MarksForeColor
  Label2.ForeColor() = New_MarksForeColor
  Label3.ForeColor() = New_MarksForeColor
  PropertyChanged "MarksForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblApp1,lblApp1,-1,Font
Public Property Get TitleFont() As Font
  Set TitleFont = m_TitleFont
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
  m_TitleFont = New_TitleFont
  Set lblApp1.Font = New_TitleFont
  Set lblApp2.Font = New_TitleFont
  PropertyChanged "TitleFont"
  Create
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblApp1,lblApp1,-1,ForeColor
Public Property Get TitleForecolor() As OLE_COLOR
Attribute TitleForecolor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  TitleForecolor = lblApp1.ForeColor
End Property

Public Property Let TitleForecolor(ByVal New_TitleForecolor As OLE_COLOR)
  lblApp1.ForeColor() = New_TitleForecolor
  PropertyChanged "TitleForecolor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblApp2,lblApp2,-1,ForeColor
Public Property Get TitleShadowColor() As OLE_COLOR
Attribute TitleShadowColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  TitleShadowColor = lblApp2.ForeColor
End Property

Public Property Let TitleShadowColor(ByVal New_TitleShadowColor As OLE_COLOR)
  lblApp2.ForeColor() = New_TitleShadowColor
  PropertyChanged "TitleShadowColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblVers,lblVers,-1,Font
Public Property Get VersionFont() As Font
Attribute VersionFont.VB_Description = "Devuelve un objeto Font."
  Set VersionFont = lblVers.Font
End Property

Public Property Set VersionFont(ByVal New_VersionFont As Font)
  Set lblVers.Font = New_VersionFont
  PropertyChanged "VersionFont"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=lblVers,lblVers,-1,ForeColor
Public Property Get VersionForeColor() As OLE_COLOR
Attribute VersionForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  VersionForeColor = lblVers.ForeColor
End Property

Public Property Let VersionForeColor(ByVal New_VersionForeColor As OLE_COLOR)
  lblVers.ForeColor() = New_VersionForeColor
  PropertyChanged "VersionForeColor"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=PicCoded,PicCoded,-1,Enabled
Public Property Get VisibleCoded() As Boolean
Attribute VisibleCoded.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
  VisibleCoded = PicCoded.Visible
End Property

Public Property Let VisibleCoded(ByVal New_VisibleCoded As Boolean)
  PicCoded.Visible = New_VisibleCoded
  PropertyChanged "VisibleCoded"
  Create
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get TitleYPos() As Integer
  TitleYPos = tYPos
End Property

Public Property Let TitleYPos(ByVal New_TitleYPos As Integer)
  tYPos = New_TitleYPos
  PropertyChanged "TitleYPos"
  Create
End Property


