VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "madgear's Definition Update"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "Process Files"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1590
      Width           =   1845
   End
   Begin VB.CommandButton cmd 
      Caption         =   "File Deleter"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   1845
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Registry Fixer"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   1845
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Virus Definition"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
VirusSig.Show 1, Me
Case 1
RegFixer.Show 1, Me
End Select
End Sub

Private Sub Form_Load()
optfile = App.Path & "\def.vfd"
End Sub
