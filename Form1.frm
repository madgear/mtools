VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "madgear v2"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   1680
      ScaleHeight     =   4755
      ScaleWidth      =   6825
      TabIndex        =   47
      Top             =   1560
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
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
         Left            =   4110
         TabIndex        =   54
         Top             =   4230
         Width           =   1275
      End
      Begin VB.ListBox memlist 
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
         Height          =   3630
         Left            =   60
         TabIndex        =   53
         Top             =   480
         Width           =   6705
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
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
         Left            =   2835
         TabIndex        =   50
         Top             =   4230
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Kill Process"
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
         Left            =   1545
         TabIndex        =   49
         Top             =   4230
         Width           =   1275
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -210
         ScaleHeight     =   615
         ScaleWidth      =   7155
         TabIndex        =   48
         Top             =   -210
         Width           =   7155
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Task Killer"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   51
            Top             =   270
            Width           =   6855
         End
      End
   End
   Begin madtools.gradient gradient8 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   43
      Top             =   8295
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   767
      GradColor1      =   12017457
      GradColor2      =   0
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   8175
         TabIndex        =   46
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © madgear ® 2007 by: Monastrial, Anthony D. | Email : madgeardx@yahoo.com | Cell : 0915-459-0695"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   210
         TabIndex        =   44
         Top             =   90
         Width           =   9660
      End
   End
   Begin madtools.gradient gradient7 
      Height          =   2745
      Left            =   8190
      TabIndex        =   31
      Top             =   420
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   4842
      GradColor1      =   192
      GradColor2      =   0
      Begin madtools.gradient gradient3 
         Height          =   465
         Index           =   3
         Left            =   -60
         TabIndex        =   32
         Top             =   -60
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   820
         GradColor1      =   0
         GradColor2      =   4210752
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   720
            TabIndex        =   39
            Top             =   150
            Width           =   525
         End
      End
      Begin VB.Label vInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#####"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label vInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#####"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   450
      End
      Begin VB.Label vInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#####"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Size"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1110
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Name :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   540
         Width           =   945
      End
   End
   Begin madtools.gradient gradient6 
      Height          =   3885
      Left            =   8190
      TabIndex        =   21
      Top             =   3150
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   6853
      GradColor1      =   192
      GradColor2      =   0
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   450
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":038A
         Caption         =   "&Start Scan"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   870
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":0724
         Caption         =   "Scan &Location"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1290
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":0ABE
         Caption         =   "&Repair"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   1710
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":0E58
         Caption         =   "&Delete"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   4
         Left            =   90
         TabIndex        =   26
         Top             =   2130
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":11F2
         Caption         =   "&Quarantine"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   5
         Left            =   90
         TabIndex        =   27
         Top             =   2550
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":158C
         Caption         =   "&Task Killer"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   6
         Left            =   90
         TabIndex        =   28
         Top             =   2970
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":1926
         Caption         =   "&About"
         BackColor       =   16777215
      End
      Begin madtools.SideMenu SideMenu1 
         Height          =   390
         Index           =   7
         Left            =   90
         TabIndex        =   29
         Top             =   3390
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         Picture         =   "Form1.frx":1CC0
         Caption         =   "&Close"
         BackColor       =   16777215
      End
      Begin madtools.gradient gradient3 
         Height          =   465
         Index           =   2
         Left            =   -30
         TabIndex        =   30
         Top             =   -60
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   820
         GradColor1      =   0
         GradColor2      =   4210752
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   690
            TabIndex        =   40
            Top             =   150
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3765
      Left            =   2100
      ScaleHeight     =   3735
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   4605
      Begin VB.ListBox sProcess 
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
         Height          =   270
         Left            =   1170
         TabIndex        =   52
         Top             =   1830
         Width           =   1365
      End
      Begin VB.FileListBox fScript 
         Appearance      =   0  'Flat
         Height          =   420
         Hidden          =   -1  'True
         Left            =   3840
         Pattern         =   "*.vbs;*.inf"
         System          =   -1  'True
         TabIndex        =   45
         Top             =   900
         Width           =   495
      End
      Begin VB.ListBox newProcess 
         Height          =   645
         Left            =   3840
         TabIndex        =   6
         Top             =   180
         Width           =   585
      End
      Begin VB.ListBox reglist 
         Height          =   450
         Left            =   1410
         TabIndex        =   5
         Top             =   660
         Width           =   525
      End
      Begin VB.ListBox virsize 
         Height          =   255
         Left            =   990
         TabIndex        =   4
         Top             =   630
         Width           =   135
      End
      Begin VB.ListBox virname 
         Height          =   255
         Left            =   1170
         TabIndex        =   3
         Top             =   630
         Width           =   135
      End
      Begin VB.ListBox deflist 
         Height          =   450
         Left            =   60
         TabIndex        =   2
         Top             =   630
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   450
         Top             =   1380
      End
      Begin VB.FileListBox File1 
         Height          =   480
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   2415
      End
   End
   Begin madtools.gradient gradient4 
      Height          =   4965
      Left            =   30
      TabIndex        =   10
      Top             =   420
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   8758
      GradColor1      =   49152
      GradColor2      =   0
      Begin VB.ListBox drivelist 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   600
         Width           =   1605
      End
      Begin madtools.gradient gradient3 
         Height          =   465
         Index           =   1
         Left            =   -150
         TabIndex        =   12
         Top             =   -30
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   820
         GradColor1      =   255
         GradColor2      =   0
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Drives"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   300
            TabIndex        =   42
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin madtools.gradient gradient2 
      Height          =   4965
      Left            =   1770
      TabIndex        =   7
      Top             =   420
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8758
      GradColor1      =   49152
      GradColor2      =   0
      Begin VB.ListBox List1 
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
         Height          =   4350
         Left            =   90
         TabIndex        =   9
         Top             =   510
         Width           =   6225
      End
      Begin madtools.gradient gradient3 
         Height          =   465
         Index           =   0
         Left            =   -210
         TabIndex        =   8
         Top             =   -30
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   820
         GradColor1      =   255
         GradColor2      =   0
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scan Details"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   390
            TabIndex        =   41
            Top             =   120
            Width           =   960
         End
      End
   End
   Begin madtools.gradient gradient5 
      Height          =   1635
      Left            =   30
      TabIndex        =   13
      Top             =   5400
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   2884
      GradColor1      =   4210752
      GradColor2      =   0
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
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   120
         Width           =   1905
      End
      Begin VB.Label virno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1260
         TabIndex        =   20
         Top             =   330
         Width           =   90
      End
      Begin VB.Label tfs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1860
         TabIndex        =   19
         Top             =   90
         Width           =   90
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Found :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Files Scanned :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   90
         Width           =   1665
      End
      Begin VB.Label Statuslbl 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   150
         TabIndex        =   16
         Top             =   690
         Width           =   7845
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Files to Scan :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4980
         TabIndex        =   15
         Top             =   180
         Width           =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   8
         X1              =   60
         X2              =   9840
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&

Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46


Dim ancho As Integer

Private Const MAX_PATH& = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim k As Integer
Dim Zoom As Integer
Dim Speed As Integer

Dim mFileSize As Long
Dim arrByte() As Byte
Dim arrSearchByte() As Byte
Dim pageStart As Long
Dim pageEnd As Long
Dim prevFoundPos As Long


Public Function KillApp(myName As String) As Boolean
Const PROCESS_ALL_ACCESS = 0
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim ExitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim i As Integer
On Local Error GoTo Finish
appCount = 0
Const TH32CS_SNAPPROCESS As Long = 2&

uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
memlist.Clear
Do While rProcessFound
    i = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    memlist.AddItem (szExename)
    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillApp = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, ExitCode)
        Call CloseHandle(myProcess)
    End If
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
Call CloseHandle(hSnapshot)
Finish:
End Function


Private Sub SearchDirs(curpath$)
On Error Resume Next
Dim dirs%, dirbuf$(), i%

File1.Path = curpath$
For no = 0 To File1.ListCount - 1
DoEvents
If FileExist(File1.Path & File1.List(no)) = True Then
tfs = Val(tfs) + 1
Scan File1.Path & File1.List(no)
End If
DoEvents
Next



hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
If hItem& <> INVALID_HANDLE_VALUE Then
Do
If (WFD.dwFileAttributes And vbDirectory) Then
If Asc(WFD.cFileName) <> vbKeyDot Then
File1.Path = curpath$ & WFD.cFileName
For no = 0 To File1.ListCount - 1
DoEvents
If FileExist(File1.Path & "\" & File1.List(no)) = True Then
tfs = Val(tfs) + 1
Scan File1.Path & "\" & File1.List(no)
End If
DoEvents
Next

If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
dirs% = dirs% + 1
dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
End If

End If
Loop While FindNextFile(hItem&, WFD)
Call FindClose(hItem&)
End If

For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
End Sub


Private Sub SearchFileSpec(curpath$)
hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
If hFile& <> INVALID_HANDLE_VALUE Then
Do
DoEvents
SendMessage hLB&, LB_ADDSTRING, 0, _
ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
MsgBox WFD.cFileName
Loop While FindNextFile(hFile&, WFD)
Call FindClose(hFile&)
End If
End Sub




Private Sub Combo2_Click()
Select Case Combo2.ListIndex
Case 0
File1.Pattern = "*.exe;*.com;*.pif;*.bat;*.scr;*.msi;*.vbs"
Case 1
File1.Pattern = "*.jpg;*.bmp;*.png;*.gif;*.pcx"
Case 2
File1.Pattern = "*.doc;*.xls;*.mdb;*.txt;*.ppt"
Case 3
File1.Pattern = "*.*"
End Select
End Sub





Private Sub Command1_Click()
If Len(memlist.List(memlist.ListIndex)) = 0 Then Exit Sub
msg = MsgBox("Are you sure you want to close this application - " & memlist.List(memlist.ListIndex), vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub
ForceKill memlist.List(memlist.ListIndex)
RefreshMemory
End Sub

Private Sub Command2_Click()
RefreshMemory
End Sub

Private Sub Command3_Click()
Picture3.Visible = False
End Sub

Private Sub drivelist_ItemCheck(Item As Integer)
'If drivelist.Selected(Item) = True Then MsgBox "ala"
End Sub



Private Sub Form_Load()
SetWindowShape Picture3.hWnd, 1

loaddrive
DefFile = App.Path & "\def.vfd"
LoadList
With File1

.Pattern = "*.exe;*.com;*.pif;*.bat;*.scr;*.msi"
.Hidden = True
.System = True
.ReadOnly = True
.Archive = True

End With
Combo2.AddItem "Executable Files"
Combo2.AddItem "Picture Files"
Combo2.AddItem "Document Files"
Combo2.AddItem "All Files"
Combo2.ListIndex = 0
MemScan

If List1.ListCount <> 0 Then
msg = MsgBox("Virus Found Active!, Remove this Virus?", vbQuestion + vbYesNo)
If msg <> vbYes Then Exit Sub
Call SideMenu1_Click(2)
End If

For no = 0 To memlist.ListCount - 1
sProcess.AddItem memlist.List(no)
Next

End Sub

Private Sub doHexSearch(HexSearch, fname As String, vname, vsize, StartSize As Long)

On Error Resume Next
Dim HexCtn As Integer
Dim i, j
Dim mMatch As Boolean
Dim foundStartPos As Long

If FileLen(fname) = 0 Then Exit Sub
foundStartPos = 0
prevFoundPos = StartSize
HexCtn = Len(HexSearch) / 2

mHandle = FreeFile
Open fname For Binary As #mHandle

mFileSize = StartSize + 10

ReDim arrByte(1 To mFileSize)
Get #mHandle, , arrByte
Close mHandle

ReDim arrhexbyte(1 To HexCtn)


For i = 1 To HexCtn
arrhexbyte(i) = CByte("&h" & (Mid(HexSearch, (i * 2 - 1), 2)))
Next i

foundStartPos = prevFoundPos + 1

For i = foundStartPos To (UBound(arrByte) - (HexCtn - 1))


If arrByte(i) = arrhexbyte(1) Then
mMatch = True

For j = 1 To (HexCtn - 1)
If arrByte(i + j) <> arrhexbyte(1 + j) Then
mMatch = False
Exit For
End If
Next j

If mMatch = True Then

ForceKill fname


List1.AddItem fname
virname.AddItem vname
virsize.AddItem vsize
virno = List2.ListCount


End If
End If
Next i
prevFoundPos = 0
End Sub


Sub Scan(sFileName As String)
On Error Resume Next
Dim defArr
DoEvents
For no = 0 To deflist.ListCount - 1
defArr = Split(deflist.List(no), ",")

Statuslbl.Caption = "Scanning... " & sFileName
doHexSearch defArr(1), sFileName, defArr(0), defArr(2), Val(defArr(3))

Next
DoEvents
End Sub

Sub loaddrive()
drvbitmask& = GetLogicalDrives()
If drvbitmask& Then
drivelist.Clear
UseFileSpec% = True
maxpwr% = Int(Log(drvbitmask&) / Log(2))
For pwr% = 0 To maxpwr%
drivelist.AddItem Chr$(vbKeyA + pwr%) & ":\"
Next
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObject Me.hWnd
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObject Picture3.hWnd
End Sub

Private Sub List1_Click()
On Error Resume Next
vInfo(0) = virname.List(List1.ListIndex)
vInfo(1) = virsize.List(List1.ListIndex)
vInfo(2) = FileLen(List1.List(List1.ListIndex))
End Sub



Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragObject Picture3.hWnd
End Sub

Private Sub SideMenu1_Click(Index As Integer)
Select Case Index
Case 0

Timer1.Enabled = False

tfs = 0
virno = 0
virname.Clear
virsize.Clear
List1.Clear

For no = 0 To drivelist.ListCount - 1

    

    If drivelist.Selected(no) = True Then
    
    
    fScript.Path = drivelist.List(no)
    fScript.Refresh
    
    If fScript.ListCount <> 0 Then
     For i = 0 To fScript.ListCount - 1
     List1.AddItem drivelist.List(no) & fScript.List(i)
     Select Case CheckExtention(drivelist.List(no) & fScript.List(i))
     Case "VBS"
     virname.AddItem "VBS Script Exploit"
     Case "INF"
     virname.AddItem "Exploit Launcher"
     End Select
     virsize.AddItem FileLen(drivelist.List(no) & fScript.List(i))
     Next
    End If
    
    
    SearchDirs drivelist.List(no)
    
    
    
    
    End If

Next

Statuslbl = "Finish Scanning!"

Case 1
Case 2


For no = 0 To List1.ListCount - 1
Statuslbl = "Cleaning file : " & List1.List(no)
VirusCheck List1.List(no), virname.List(no), virsize.List(no)
Next

Statuslbl = "Finish cleaning files..."
List1.Clear

Case 3
'delete
Case 4
'quaratine
Case 5
Picture3.Visible = True
'taskkiller
Case 6
'about
Case 7
msg = MsgBox("Are you sure you want to close this program?", vbQuestion + vbYesNo, Label1)
If msg <> vbYes Then Exit Sub
End
End Select
End Sub

Private Sub SideMenu1_MouseExit(Index As Integer)
SideMenu1(Index).BackColor = vbWhite
End Sub

Private Sub SideMenu1_MouseOver(Index As Integer)
SideMenu1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Timer1_Timer()
If Picture3.Visible = True Then Timer1.Enabled = False
RefreshMemory
RemoveExcess sProcess, memlist, newProcess

If newProcess.ListCount <> 0 Then
For no = 0 To newProcess.ListCount - 1
Scan newProcess.List(no)
Next
Statuslbl = "Finish Scanning!"
End If

If List1.ListCount <> 0 Then
Call SideMenu1_Click(2)
Else
For no = 0 To newProcess.ListCount - 1
sProcess.AddItem newProcess.List(no)
Next
End If

If memlist.ListCount < sProcess.ListCount Then
sProcess.Clear
For no = 0 To memlist.ListCount - 1
sProcess.AddItem memlist.List(no)
Next
End If

End Sub

Sub LoadList()

defnum = Val(GetINIFile(DefFile, "def", "no"))
regnum = Val(GetINIFile(DefFile, "regfix", "no"))

For no = 1 To defnum
deflist.AddItem GetINIFile(DefFile, "def", "vir" & CStr(no))
Next
For no = 1 To regnum
reglist.AddItem GetINIFile(DefFile, "regfix", "entry" & CStr(no))
Next


End Sub

Sub MemScan()
Statuslbl.Visible = False
Timer1.Enabled = False

RefreshMemory

For no = 0 To memlist.ListCount - 1
DoEvents
Scan memlist.List(no)
DoEvents
Next


Timer1.Enabled = True

Statuslbl = ""
Statuslbl.Visible = True
End Sub

Sub RefreshMemory()

'On Error Resume Next

memlist.Clear


If CheckOS = "WinNT" Then

Dim processes As New ProcessList
For i = 0 To processes.processCount - 1
memlist.AddItem processes.ProcessName(i)
Next
Set processes = Nothing

Else

KillApp ("none")

End If

End Sub

Function CheckExtention(xt As String) As String
CheckExtention = UCase(Right(xt, 3))
End Function
