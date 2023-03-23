VERSION 5.00
Begin VB.Form fMinShow 
   BackColor       =   &H000000A0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2775
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "WEIGHTSMART"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   3075
   End
End
Attribute VB_Name = "fMinShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Check As Integer

Public Sub Form_Load()
Dim lreigon
    lblTitle.Width = Me.Width
    lblTitle.Left = 0
    'Green Shade
    'lblTitle.BackColor = RGB(100, 200, 150)
    'Purple Shade
    Me.Left = Screen.Width - (Me.Width)
    Me.Top = Screen.Height
    Me.Show
    'Me.AutoRedraw = True
    Timer.Enabled = True
    Check = 0
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub Timer_Timer()
If Not Check = 300 Then
    If Not Me.Top = Screen.Height - (Me.Height + 400) Then
        Me.Top = Me.Top - 100
    End If
    Check = Check + 1
Else
        If Not Screen.Height = Me.Top Then
            Me.Top = Me.Top + 100
        Else
            Timer.Enabled = False
            Unload Me
        End If
End If
End Sub

