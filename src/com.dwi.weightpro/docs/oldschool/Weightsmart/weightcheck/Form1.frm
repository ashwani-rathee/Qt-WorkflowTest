VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox comwt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Text            =   "2"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Port"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   420
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Text            =   "2400,n,8,1"
      Top             =   420
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   240
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = False
MSComm1.PortOpen = False
MSComm1.Settings = Text1.Text
MSComm1.CommPort = Text2.Text
MSComm1.PortOpen = True
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
MSComm1.PortOpen = False
If MSComm1.CommPort = 2 Then
    MSComm1.CommPort = 1
    Text2.Text = "1"
Else
    MSComm1.CommPort = 2
    Text2.Text = "2"
End If
MSComm1.PortOpen = True
Command1_Click
End Sub

Private Sub Form_Load()
MSComm1.CommPort = 1
MSComm1.PortOpen = True
End Sub



Private Sub Timer1_Timer()
    a = MSComm1.Input
'    For i = 1 To Len(a) - 6
'        If IsNumeric(Mid(a, i, 6)) Then
'            b = CLng(Mid(a, i, 6))
'            comwt.Caption = b
'            Exit For
'        End If
'    Next
comwt.Text = a
End Sub
