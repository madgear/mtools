Attribute VB_Name = "shape"



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
            hRgn = CreateRoundRectRgn(0, 0, wi, he, 20, 20)
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








