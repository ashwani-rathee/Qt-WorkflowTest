VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BCCL HEADQUARTER"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   18465
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   19020
      TabIndex        =   3
      Top             =   1635
      Width           =   19080
      Begin VB.CommandButton Command7 
         Caption         =   "WT. DIFF. EXCEPTIONS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   14
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "USERS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12600
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "TAGS ISSUED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "VEHICLE TRACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SUMMARY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16440
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TAG EXCEPTIONS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "WEIGHMENT REPORTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H000000C0&
      ForeColor       =   &H000000C0&
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   19020
      TabIndex        =   0
      Top             =   0
      Width           =   19080
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   13920
         TabIndex        =   11
         Top             =   120
         Width           =   1815
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   1680
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "DADHWAL"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   800
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "INSTALLED AND MAINTAINED BY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   120
            TabIndex        =   12
            Top             =   160
            Width           =   1815
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HEADQUARTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   11415
      End
      Begin VB.Image Image1 
         Height          =   1530
         Left            =   360
         Picture         =   "MDIForm1.frx":2D7A3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reports for RFID Vehicle Access Control System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BHARAT COKING COAL LIMITED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   1
         Top             =   0
         Width           =   11415
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub unloadall()
Unload Form1
Unload exepreport
Unload tagtrack
Unload tagissued
Unload Users
End Sub

Private Sub Command1_Click()
unloadall
Form1.Show
End Sub

Private Sub Command2_Click()
unloadall
exepreport.Show
End Sub

Private Sub Command3_Click()
unloadall
sumreport.Show
End Sub

Private Sub Command4_Click()
unloadall
tagtrack.Show
End Sub

Private Sub Command5_Click()
unloadall
tagissued.Show
End Sub

Private Sub Command6_Click()
unloadall
Users.Show
End Sub

Private Sub Command7_Click()
unloadall
weightdiff.Show
End Sub
