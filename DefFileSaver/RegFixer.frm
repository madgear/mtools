VERSION 5.00
Begin VB.Form RegFixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry Fixer"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox regList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   90
      TabIndex        =   17
      Top             =   60
      Width           =   7245
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   90
      ScaleHeight     =   3765
      ScaleWidth      =   7215
      TabIndex        =   2
      Top             =   2610
      Width           =   7245
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   10
         Top             =   90
         Width           =   3345
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   9
         Top             =   510
         Width           =   2565
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   8
         Top             =   930
         Width           =   2025
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   7
         Top             =   1710
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   6
         Top             =   2130
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   2550
         Width           =   3975
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3210
         Width           =   1635
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3210
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg Entry :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   60
         TabIndex        =   15
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   60
         TabIndex        =   14
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   60
         TabIndex        =   13
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   60
         TabIndex        =   12
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "String :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   60
         TabIndex        =   11
         Top             =   2580
         Width           =   660
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2100
      Width           =   1635
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      Width           =   1635
   End
End
Attribute VB_Name = "RegFixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub loadlist()
On Error GoTo errhand
regnum = Val(GetINIFile(optfile, "regfix", "no"))
rNum = regnum
For no = 1 To regnum
regList.AddItem GetINIFile(optfile, "regfix", "entry" & CStr(no))
Next
errhand:
End Sub

Private Sub cmd_Click(Index As Integer)
On Error GoTo errhand
Select Case Index

Case 0

For no = 0 To regList.ListCount - 1
WriteIni optfile, "regfix", "entry" & CStr(no + 1), (regList.List(no))
Next
WriteIni optfile, "regfix", "no", regList.ListCount

Case 1

regList.RemoveItem regList.ListIndex

Case 2
If Combo2 = "SET" Then
regList.AddItem Text1 & "," & Text2 & "," & Combo1.Text & "," & "1" & "," & Combo3.Text & "," & Text3
Else
regList.AddItem Text1 & "," & Text2 & "," & Combo1.Text & "," & "0"
End If
Case 3
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Text1 = ""
Text2 = ""
Text3 = ""
Combo1.SetFocus
End Select
errhand:
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "SET" Then
Combo3.Enabled = True
Text3.Enabled = True
Else
Combo3.Enabled = False
Text3.Enabled = False
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()

loadlist
Combo1.AddItem "HKLM"
Combo1.AddItem "HKCU"
Combo1.AddItem "HCR"
Combo1.AddItem "HU"
Combo2.AddItem "DELETE"
Combo2.AddItem "SET"
Combo3.AddItem "STRING"
Combo3.AddItem "DWORD"

End Sub



