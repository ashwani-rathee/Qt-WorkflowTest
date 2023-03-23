VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00FDF2FD&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5295
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   8070
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fSplash.frx":030A
   ScaleHeight     =   5295
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "WeighBridge"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Computerised "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000A0&
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Global Infotec "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software.  Protected by Global Infotec"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   3810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   4200
      TabIndex        =   3
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by:- "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1200
      TabIndex        =   1
      Top             =   4530
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All  Rights  are  reserved.  Don't  make  illegal  copies of this "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1200
      TabIndex        =   0
      Top             =   4800
      Width           =   6375
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Varify()

End Sub

Private Sub Form_Load()
    Me.Height = 5295
    Me.Width = 7245
    Me.Show
    
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Dim ltime
    ltime = Timer()
    While Timer() - ltime < 1
        'Do Nothing
    Wend
'    Call Varify
'    Call TrialOver
    'Call ConnectComp
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    fMinShow.Form_Load
    'fLogin.Show 1
    fLogin.Show
End Sub
