VERSION 5.00
Begin VB.UserControl gradient 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape gr2 
      FillColor       =   &H00B75F31&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   330
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape gr1 
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00F0C797&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1560
      Top             =   150
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim grad1 As OLE_COLOR
Dim grad2 As OLE_COLOR


Public Red, Green, Blue

Public Enum GradStyle
    Vertical = 0
    Horizontal = 1
    SideRight = 2
    SideLeft = 3
End Enum



Dim gStyle As GradStyle

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Region API functins
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, _
    ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
    lpRect As RECT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
    ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As _
    Long
Function ApplyGradient(Way As String, SClr As ColorConstants, EClr As ColorConstants)

Dim SRed, ERed, SGreen, SBlue, EGreen, EBlue
Dim DifR, DifG, DifB
Dim Fora
Dim Yi

UserControl.AutoRedraw = True: UserControl.ScaleMode = 3
Analyze (SClr): SRed = Red: SGreen = Green: SBlue = Blue
Analyze (EClr): ERed = Red: EGreen = Green: EBlue = Blue
DifR = ERed - SRed: DifG = EGreen - SGreen: DifB = EBlue - SBlue
Select Case UCase(Way)
  Case Is = "H": Fora = UserControl.ScaleHeight
  Case Is = "V": Fora = UserControl.ScaleWidth
  Case "DU", "DD": Fora = Sqr(2) * Sqr((UserControl.ScaleWidth * UserControl.ScaleWidth) + (UserControl.ScaleHeight * UserControl.ScaleHeight))
End Select
For Yi = 0 To Fora
SRed = SRed + (DifR / Fora): If SRed < 0 Then SRed = 0
SGreen = SGreen + (DifG / Fora): If SGreen < 0 Then SGreen = 0
SBlue = SBlue + (DifB / Fora): If SBlue < 0 Then SBlue = 0
Select Case UCase(Way)
  Case Is = "H": UserControl.Line (0, Yi)-(UserControl.ScaleWidth, Yi), RGB(SRed, SGreen, SBlue), B
  Case Is = "V": UserControl.Line (Yi, 0)-(Yi, UserControl.ScaleHeight), RGB(SRed, SGreen, SBlue), B
  Case Is = "DU": UserControl.Line (Yi, 0)-(0, Yi), RGB(SRed, SGreen, SBlue)
  Case Is = "DD": UserControl.Line (UserControl.ScaleWidth - Yi, 0)-(UserControl.ScaleWidth, Yi), RGB(SRed, SGreen, SBlue)
End Select
Next

UserControl.AutoRedraw = True: UserControl.ScaleMode = 1
End Function
Function Analyze(Chroma As ColorConstants)
Dim rr, gr, br As Long
rr = 1: gr = 256: br = 65536
Dim rest As Long
rest = Chroma \ br
Blue = rest
Chroma = Chroma Mod br
If Blue < 0 Then Blue = 0
rest = Chroma \ gr
Green = rest
Chroma = Chroma Mod gr
If Green < 0 Then Green = 0
rest = Chroma \ rr
Red = rest
Chroma = Chroma Mod rr
If Red < 0 Then Red = 0
End Function

Private Sub UserControl_Initialize()
gr1.FillColor = GradColor1
gr2.FillColor = GradColor2
repos
End Sub



Private Sub UserControl_Resize()
repos
End Sub

Sub repos()
gradapply
'Call ApplyGradient("H", &HB75F31, &HF0C797)
SetWindowShape hWnd, 1

End Sub




' modify the shape of a window
'
' This routine supports three values for SHAPE
'    0 = circle/ellipse, 1=rounded rect, 2=rhomb
'
' NOTES: You get best effects using borderless forms
'        Remember to provide alternative commands for
'        closing and moving the form

Sub SetWindowShape(ByVal hWnd As Long, ByVal shape As Long)
    Dim lpRect As RECT
    Dim wi As Long, he As Long
    Dim hRgn As Long
    
    ' get the bounding rectangle's size
    GetWindowRect hWnd, lpRect
    wi = lpRect.Right - lpRect.Left
    he = lpRect.Bottom - lpRect.Top
    
    ' create a region
    Select Case shape
        Case 0          ' circle/ellipse
            hRgn = CreateEllipticRgn(0, 0, wi, he)
        Case 1          ' rounded rectangle
            hRgn = CreateRoundRectRgn(2, 2, wi, he, 12, 12)
        Case 2          ' rhomb
            Dim lpPoints(3) As POINTAPI
            lpPoints(0).X = wi \ 2
            lpPoints(0).Y = 0
            lpPoints(1).X = 0
            lpPoints(1).Y = he \ 2
            lpPoints(2).X = wi \ 2
            lpPoints(2).Y = he
            lpPoints(3).X = wi
            lpPoints(3).Y = he \ 2
            hRgn = CreatePolygonRgn(lpPoints(0), 4, 1)
    End Select
    
    ' trim the window to the region
    SetWindowRgn hWnd, hRgn, True
    DeleteObject hRgn

End Sub




Public Property Get GradColor1() As OLE_COLOR
On Error Resume Next
    GradColor1 = gr1.FillColor
    
End Property
Public Property Let GradColor1(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    gr1.FillColor = New_BackColor
    PropertyChanged "GradColor1"
    gradapply
    Refresh
End Property

Public Property Get GradColor2() As OLE_COLOR
On Error Resume Next
    GradColor2 = gr2.FillColor
  
End Property
Public Property Let GradColor2(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    gr2.FillColor = New_BackColor
    PropertyChanged "GradColor2"
    gradapply
    Refresh
End Property


Public Property Get GradientStyle() As GradStyle
    GradientStyle = gStyle
    
End Property
Public Property Let GradientStyle(ByVal New_PicturePosition As GradStyle)
    gStyle = New_PicturePosition
    PropertyChanged "GradientStyle"
    gradapply
    Refresh
End Property



Private Sub UserControl_Show()
repos
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("GradColor1", gr1.FillColor, &H8000000F)
Call PropBag.WriteProperty("GradColor2", gr2.FillColor, &H8000000F)
Call PropBag.WriteProperty("GradientStyle", gStyle, 1)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
gr1.FillColor = PropBag.ReadProperty("GradColor1", &H8000000F)
gr2.FillColor = PropBag.ReadProperty("GradColor2", &H8000000F)
gStyle = PropBag.ReadProperty("GradientStyle", 1)
End Sub

Sub gradapply()
Dim dum1 As String
Select Case gStyle
Case 0
dum1 = "V"
Case 1
dum1 = "H"
Case 2
dum1 = "DD"
Case 3
dum1 = "DU"
End Select
Call ApplyGradient(dum1, gr1.FillColor, gr2.FillColor)
End Sub
