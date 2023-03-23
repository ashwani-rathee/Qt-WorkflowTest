VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form comm 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3840
      Top             =   3840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1680
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   1200
      DataBits        =   7
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   2655
   End
End
Attribute VB_Name = "comm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim a As String
Dim i As Integer

MSComm1.CommPort = 1
MSComm1.PortOpen = True
End Sub

Private Sub Timer1_Timer()
a = MSComm1.Input
a = Mid(a, 1, 7)
If Trim(a) <> "" Then
Label1 = a
End If


End Sub
