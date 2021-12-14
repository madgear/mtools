VERSION 5.00
Begin VB.UserControl SideMenu 
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1590
   ScaleWidth      =   4800
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   240
   End
   Begin VB.Label c 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SideMenu"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   390
      TabIndex        =   0
      Top             =   75
      Width           =   780
   End
   Begin VB.Image i 
      Enabled         =   0   'False
      Height          =   240
      Left            =   60
      Stretch         =   -1  'True
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "sidemenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim m_Picture As Picture

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Const g_Light = &H80000016
Const g_Shadow = &H80000010
Const g_HighLight = &H80000014
Const g_DarkShadow = &H80000015


Event MouseOver()
Event MouseExit()
Event Click()

Private Sub c_Click()
RaiseEvent Click
End Sub

Private Sub c_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlClick
End Sub

Private Sub c_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlUp
End Sub

Private Sub i_Click()
RaiseEvent Click
End Sub

Private Sub i_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlClick
End Sub

Private Sub i_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlUp
End Sub

Private Sub OverTimer_Timer()

    Dim P As POINTAPI
    GetCursorPos P
    
    If UserControl.hWnd <> WindowFromPoint(P.X, P.Y) Then
    RaiseEvent MouseExit
    OverTimer.Enabled = False
    End If
    
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub




Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlClick
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
  
 
  RaiseEvent MouseOver
  OverTimer.Enabled = True
  
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ctrlUp
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set i.Picture = PropBag.ReadProperty("Picture", Nothing)
    c.Caption = PropBag.ReadProperty("Caption", "SideMenu")
 
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Height = 390

With i
.Top = 75
.Left = 60
End With

With c
.Top = 75
.Left = 390
End With

End Sub
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set i.Picture = New_Picture
    PropertyChanged "Picture"
    Refresh
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Picture", i.Picture, Nothing)
  Call PropBag.WriteProperty("Caption", c.Caption, "SideMenu")
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
On Error Resume Next
    Caption = c.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
On Error Resume Next
    c.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Sub ctrlClick()
i.Top = 85
i.Left = 70
c.Top = 85
c.Left = 400
End Sub
Sub ctrlUp()
i.Top = 75
i.Left = 60
c.Top = 75
c.Left = 390
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property


Sub Refresh()

If Enabled = False Then
c.ForeColor = &H80000011
Else
c.ForeColor = vbBlack
End If
End Sub
