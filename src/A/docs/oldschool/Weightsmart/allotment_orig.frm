VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form allotment 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DO Allotment"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14475
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   14475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   111
      ToolTipText     =   "Unload Form"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   11160
      TabIndex        =   98
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   11160
      TabIndex        =   97
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   11160
      TabIndex        =   96
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   11160
      TabIndex        =   95
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   11160
      TabIndex        =   94
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker tmpdt 
      Height          =   255
      Left            =   120
      TabIndex        =   93
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   81920001
      CurrentDate     =   41165
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1200
      TabIndex        =   47
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
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
      Left            =   11640
      TabIndex        =   46
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   9840
      TabIndex        =   45
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   11160
      TabIndex        =   44
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   11160
      TabIndex        =   43
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   11160
      TabIndex        =   42
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   11160
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   11160
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   11160
      TabIndex        =   39
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   11160
      TabIndex        =   38
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   11160
      TabIndex        =   37
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   11160
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   11160
      TabIndex        =   35
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   11160
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   6840
      TabIndex        =   33
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   6840
      TabIndex        =   32
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   6840
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   6840
      TabIndex        =   30
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   6840
      TabIndex        =   29
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   6840
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   6840
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   6840
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   6840
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   6840
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   6840
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   6840
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   6840
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   6840
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   6840
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   6840
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   6840
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2520
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   2520
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2520
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2520
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   2520
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Texta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   48
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   49
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   51
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   52
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   6
      Left            =   1200
      TabIndex        =   54
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   7
      Left            =   1200
      TabIndex        =   55
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   56
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   9
      Left            =   1200
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   10
      Left            =   1200
      TabIndex        =   58
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   60
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   13
      Left            =   1200
      TabIndex        =   61
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   14
      Left            =   1200
      TabIndex        =   62
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   15
      Left            =   1200
      TabIndex        =   63
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   16
      Left            =   1200
      TabIndex        =   64
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   17
      Left            =   5520
      TabIndex        =   65
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   18
      Left            =   5520
      TabIndex        =   66
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   19
      Left            =   5520
      TabIndex        =   67
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   20
      Left            =   5520
      TabIndex        =   68
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   21
      Left            =   5520
      TabIndex        =   69
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   22
      Left            =   5520
      TabIndex        =   70
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   23
      Left            =   5520
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   24
      Left            =   5520
      TabIndex        =   72
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   25
      Left            =   5520
      TabIndex        =   73
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   26
      Left            =   5520
      TabIndex        =   74
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   27
      Left            =   5520
      TabIndex        =   75
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   28
      Left            =   5520
      TabIndex        =   76
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   29
      Left            =   5520
      TabIndex        =   77
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   30
      Left            =   5520
      TabIndex        =   78
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   31
      Left            =   5520
      TabIndex        =   79
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   32
      Left            =   5520
      TabIndex        =   80
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   33
      Left            =   5520
      TabIndex        =   81
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   34
      Left            =   9840
      TabIndex        =   82
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   35
      Left            =   9840
      TabIndex        =   83
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   36
      Left            =   9840
      TabIndex        =   84
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   37
      Left            =   9840
      TabIndex        =   85
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   38
      Left            =   9840
      TabIndex        =   86
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   39
      Left            =   9840
      TabIndex        =   87
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   40
      Left            =   9840
      TabIndex        =   88
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   41
      Left            =   9840
      TabIndex        =   89
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   42
      Left            =   9840
      TabIndex        =   90
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   43
      Left            =   9840
      TabIndex        =   91
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   44
      Left            =   9840
      TabIndex        =   92
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   45
      Left            =   9840
      TabIndex        =   99
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   46
      Left            =   9840
      TabIndex        =   100
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   47
      Left            =   9840
      TabIndex        =   101
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   48
      Left            =   9840
      TabIndex        =   102
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin MSComCtl2.DTPicker Labela 
      Height          =   375
      Index           =   49
      Left            =   9840
      TabIndex        =   103
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81920001
      CurrentDate     =   41154
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   960
      TabIndex        =   104
      Top             =   720
      Width           =   12735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   106
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5880
         TabIndex        =   105
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81920001
         CurrentDate     =   41154
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   10200
         TabIndex        =   107
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81920001
         CurrentDate     =   41154
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DO Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   109
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   108
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Allotment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   6360
      TabIndex        =   112
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808000&
      Height          =   615
      Left            =   960
      TabIndex        =   113
      Top             =   120
      Width           =   12735
   End
End
Attribute VB_Name = "allotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Command1_Click()
Set rs3 = New ADODB.Recordset
rs3.Open "delete * from allotment where do_no='" + Trim(Text1.Text) + "'", co, adOpenKeyset, adLockOptimistic


rs3.Open "select * from allotment where do_no='" + Trim(Text1.Text) + "'", co, adOpenKeyset, adLockOptimistic
For i = 0 To n - 1
    rs3.AddNew
    rs3.Fields("do_no").Value = Trim(Text1.Text)
    rs3.Fields("w_date").Value = Labela(i).Value
    rs3.Fields("allotment").Value = Val(Trim(Texta(i).Text))
    rs3.Update
Next i
rs3.Close
MsgBox "Allotment data saved"
End Sub

Private Sub Command2_Click()
hidetext
End Sub



Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
DTPicker2.Value = DTPicker1.Value
End Sub

Sub showtext()
    n = DTPicker2.Value - DTPicker1.Value + 1
    DTPicker3.Value = DTPicker1.Value
    For i = 0 To n - 1
        Labela(i).Visible = True
        Labela(i).Value = DTPicker3.Value
        DTPicker3.Value = DTPicker3.Value + 1
        Texta(i).Visible = True
        Texta(i).Text = 0
    Next i
End Sub

Sub hidetext()
For i = 0 To 49
    Labela(i).Visible = False
    Texta(i).Visible = False
Next i
n = 0
End Sub

Private Sub DTPicker2_Change()
If DTPicker2.Value >= DTPicker1.Value Then
    If DTPicker2.Value - DTPicker1.Value <= 46 Then
        hidetext
        showtext
    Else
        DTPicker2.Value = DTPicker1.Value
        MsgBox "To date cannot be more than 45 days later than from date, Please select correct date range"
    End If
Else
    DTPicker2.Value = DTPicker1.Value
    MsgBox "Please select correct date range"
End If
End Sub

Private Sub DTPicker2_Click()
If DTPicker2.Value >= DTPicker1.Value Then
    If DTPicker2.Value - DTPicker1.Value <= 46 Then
        hidetext
        showtext
    Else
        DTPicker2.Value = DTPicker1.Value
        MsgBox "To date cannot be more than 45 days later than from date, Please select correct date range"
    End If
Else
    DTPicker2.Value = DTPicker1.Value
    MsgBox "Please select correct date range"
End If
End Sub




Private Sub Form_Load()
Me.Picture = main.Picture
Call Conn
End Sub

Sub showallot()
Set rs1 = New ADODB.Recordset
Dim i As Integer
rs1.Open "select * from allotment where do_no='" + Trim(Text1.Text) + "' order by w_date", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    i = 0
    rs1.MoveFirst
    While rs1.EOF = False
        Labela(i).Value = rs1.Fields("w_date")
        Texta(i).Text = rs1.Fields("allotment").Value
        rs1.MoveNext
        i = i + 1
    Wend
End If
rs1.Close
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If Trim(Text1.Text) <> "" And KeyAscii = 13 Then
    Set rs3 = New ADODB.Recordset
    rs3.Open "Select * from cadata where do_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
        tmpdt.Year = Mid(rs3.Fields("do_start_date").Value, 1, 4)
        tmpdt.Month = Mid(rs3.Fields("do_start_date").Value, 5, 2)
        tmpdt.Day = Mid(rs3.Fields("do_start_date").Value, 7, 2)
        DTPicker1.Value = tmpdt.Value
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
        tmpdt.Year = Mid(rs3.Fields("do_end_date").Value, 1, 4)
        tmpdt.Month = Mid(rs3.Fields("do_end_date").Value, 5, 2)
        tmpdt.Day = Mid(rs3.Fields("do_end_date").Value, 7, 2)
        DTPicker2.Value = tmpdt.Value
        showtext
        showallot
    Else
        hidetext
        MsgBox "This code does not exist in Master", vbInformation, "Code Not Found"
        Exit Sub
    End If
End If
End Sub
