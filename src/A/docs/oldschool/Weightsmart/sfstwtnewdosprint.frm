VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form sfstwt 
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   13425
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   7800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   7800
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   840
      Top             =   7800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   7800
   End
   Begin VB.CommandButton Command4 
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
      Left            =   12360
      TabIndex        =   28
      ToolTipText     =   "Unload Form"
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   2160
      TabIndex        =   23
      Top             =   6840
      Width           =   10695
      Begin VB.CommandButton cmdref 
         Caption         =   "REFRESH"
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
         Left            =   5760
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CANCEL TRANSACTION"
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
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "SAVE"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   10695
      Begin VB.TextBox Text18 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Texttag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Read Weight"
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
         Left            =   3720
         TabIndex        =   38
         Top             =   5640
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73793537
         CurrentDate     =   40639
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   3960
         TabIndex        =   34
         Top             =   3000
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   3840
         TabIndex        =   33
         Top             =   4680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   5160
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   21
         Top             =   3960
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Tag Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   50
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Shape door1 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   9000
         Top             =   5640
         Width           =   375
      End
      Begin VB.Shape door2 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   9960
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "D-1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9000
         TabIndex        =   45
         Top             =   5340
         Width           =   375
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "D-2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9960
         TabIndex        =   44
         Top             =   5340
         Width           =   375
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   5520
         TabIndex        =   43
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label d11 
         Caption         =   "0"
         Height          =   315
         Left            =   7440
         TabIndex        =   42
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label d21 
         Caption         =   "0"
         Height          =   315
         Left            =   7920
         TabIndex        =   41
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label doorin1 
         Caption         =   "0"
         Height          =   195
         Left            =   7440
         TabIndex        =   40
         Top             =   5460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "RLW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   39
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label Label16 
         Caption         =   "Press F1 For Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   36
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Press F1 For Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Party Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   32
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Material Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   31
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "First Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   27
         Top             =   5760
         Width           =   1245
      End
      Begin VB.Label Label10 
         Caption         =   "Material Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Challan No"
         Height          =   195
         Left            =   6960
         TabIndex        =   16
         Top             =   4920
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Add/Dest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   15
         Top             =   3990
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   14
         Top             =   3525
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   13
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Operator Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   12
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Daily Serial No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   9
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Session"
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
         Left            =   7920
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Simple First Weight "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   5640
      TabIndex        =   29
      Top             =   240
      Width           =   2820
   End
   Begin VB.Label Label13 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   2160
      TabIndex        =   30
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "sfstwt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim str1 As String
Dim num As Integer
Dim recfound As Boolean
Dim pressf1 As Boolean
Dim g As Integer
Dim itmfound As Boolean

Dim doorin As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim tagvalid As Boolean
Dim wtmode As String
Public OutBufferCS, nBytesWrite As Byte
Dim TXBuffer(200) As Byte
Dim RXBuffer(100) As Byte
Dim FunctionValue As Byte
Dim waitAnswerTime As Integer
Dim lastCmd As Byte
Dim tctr As Integer


Const Loc_Begin = 0
Const Loc_Temp = 1
Const Loc_Command = 2
Const Loc_Address = 3
Const Loc_DoorAddr = 4
Const Loc_Len = 5
Const Loc_Data = 7

Const STX = 2
Const ETX = 3
Const ACK = 6
Const DLE = &H10&
Const NAK = &H15&
Const SYN = &H16&


Const T_CMD_GETCARDEVENT = &H30&
Const T_CMD_GETSTATUS = &H40&
Const T_CMD_SETDOORPARA = &H61&
Const T_CMD_SETDOORDATA = &H60&
Const T_CMD_SETCONTROLPARA = &H63&
Const T_CMD_CardStatus = &H52&
Const T_CMD_AskCardEvent = &H53&
Const T_CMD_AskAlarmEvent = &H54&
Const T_CMD_GetStatusData = &H56&
Const T_CMD_GETALARMEVENT = &H3A&

Dim i As Integer

Private Function WaitAnswer()
    WaitAnswer = False
    
    waitAnswerTime = 0
    
    While (waitAnswerTime < 500)
        
        waitAnswerTime = waitAnswerTime + 1
        Sleep 5
        DoEvents
        
        If (waitAnswerTime = 0) Then
            WaitAnswer = True
            Exit Function
        End If
    Wend
    
End Function

Private Sub PutBuf(AData As Byte)
   TXBuffer(nBytesWrite) = AData
   OutBufferCS = AData Xor OutBufferCS
   nBytesWrite = nBytesWrite + 1
End Sub

Private Sub SetBufCommand(ByVal ACommand As Byte)
    TXBuffer(Loc_Begin) = (STX)
    OutBufferCS = (STX)
    nBytesWrite = Loc_Data
    TXBuffer(Loc_Command) = ACommand
    TXBuffer(Loc_DoorAddr) = 0
    TXBuffer(Loc_Len) = 0
    TXBuffer(Loc_Len + 1) = 0
   ' PutBuf (255)
End Sub



Private Sub DoSendData()
    Dim i As Integer
    Dim TXBff()   As Byte
    Dim DataLen, RN As Integer
    Dim vData, Command, s, s1 As String
    
    DataLen = nBytesWrite - Loc_Data
    TXBuffer(Loc_Len) = DataLen And 255
    TXBuffer(Loc_Len + 1) = (DataLen / 256)
    TXBuffer(Loc_Temp) = 255
    OutBufferCS = 0
    
    For i = 0 To nBytesWrite - 1
       OutBufferCS = OutBufferCS Xor TXBuffer(i)
    Next i
    
    TXBuffer(nBytesWrite) = OutBufferCS
    TXBuffer(nBytesWrite + 1) = ETX
    
    ReDim TXBff(nBytesWrite + 1)
    
    s = ""
    For i = 0 To nBytesWrite + 1
       'vData = vData + Str(TXBuffer(i))
       TXBff(i) = TXBuffer(i)
       s1 = Hex(TXBuffer(i))
        If Len(s1) = 1 Then
            s1 = "0" + s1
        End If
        s = s + s1 + " "
    Next i

    Winsock1.SendData TXBff
    Addlog "Send:  ", s
End Sub

Private Sub Addlog(ByVal titl As String, ByVal msg As String)
    'Text3.Text = Text3.Text + (Chr(13)) + Chr(10) + titl + msg
End Sub


Private Sub DoSendData1()
    Dim i As Integer
    Dim TXBff()   As Byte
    Dim DataLen, RN As Integer
    Dim vData, Command, s, s1 As String
    
    DataLen = nBytesWrite - Loc_Data
    TXBuffer(Loc_Len) = DataLen And 255
    TXBuffer(Loc_Len + 1) = (DataLen / 256)
    TXBuffer(Loc_Temp) = 255
    OutBufferCS = 0
    
    For i = 0 To nBytesWrite - 1
       OutBufferCS = OutBufferCS Xor TXBuffer(i)
    Next i
    
    TXBuffer(nBytesWrite) = OutBufferCS
    TXBuffer(nBytesWrite + 1) = ETX
    
    ReDim TXBff(nBytesWrite + 1)
    
    s = ""
    For i = 0 To nBytesWrite + 1
       'vData = vData + Str(TXBuffer(i))
       TXBff(i) = TXBuffer(i)
       s1 = Hex(TXBuffer(i))
        If Len(s1) = 1 Then
            s1 = "0" + s1
        End If
        s = s + s1 + " "
    Next i
    
    Winsock2.SendData TXBff
    Addlog "Send:  ", s
End Sub

Private Sub openclose()
    Dim re As Boolean
    Dim DataLen, Door, i As Integer
    If boomcode = 1 Then
        Door = 1
        SetBufCommand (&H2C&)
        PutBuf (Door)
        TXBuffer(Loc_DoorAddr) = Door
        checksock1
        DoSendData
    End If
End Sub

Private Sub openclose1()
    Dim re As Boolean
    Dim DataLen, Door, i As Integer
    If boomcode = 1 Then
        Door = 1
        SetBufCommand (&H2C&)
        PutBuf (Door)
        TXBuffer(Loc_DoorAddr) = Door
        checksock2
        DoSendData1
    End If
End Sub

Private Sub getTag()
Dim OutData(30000) As Byte
Dim TagCounter(20) As Byte
Dim str As String

a = Inventory(1, OutData(0), TagCounter(0))

TagNum = TagCounter(0)
If TagNum > 0 Then
    a = CleanInventory()
    For j = 0 To 5
        For i = 0 To OutData(30 * j + 6) - 1
        str = str + Right$("00" & Hex$(OutData(30 * j + 7 + i)), 2)
        Next i
    If Trim(str) <> "" Then
    Texttag.Text = str
    If Trim(Texttag) = Trim(main.texttagg) Then
        doorin = 1
        tagvalid = True
        showtagdata
    Exit Sub
    End If
    End If
    Next j
End If
End Sub

Private Sub showtagdata()
Timer2.Enabled = False
Set rs4 = New ADODB.Recordset
rs4.Open "Select * from tags where tagno = '" + Trim(Texttag) + "' and valid='1'", co1, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
rs4.MoveFirst
Text12.Text = Trim(Texttag.Text)
wtmode = rs4.Fields("mode").Value

If doorin = 1 Then
    openclose
    d1 = 1
    door1.FillColor = &HFF00&
ElseIf doorin = 2 Then
    openclose1
    d2 = 1
    door2.FillColor = &HFF00&
End If

End If
rs4.Close

End Sub

Private Sub getTagg()
Dim OutData(30000) As Byte
Dim TagCounter(20) As Byte
Dim str As String

a = Inventory(1, OutData(0), TagCounter(0))

TagNum = TagCounter(0)
If TagNum > 0 Then
    a = CleanInventory()
    For j = 0 To 5
        For i = 0 To OutData(30 * j + 6) - 1
        str = str + Right$("00" & Hex$(OutData(30 * j + 7 + i)), 2)
        Next i
    If Trim(str) <> "" Then
    Texttag.Text = str
    If Trim(Texttag) = Trim(main.texttagg) Then
        doorin = 2
        tagvalid = True
        showtagdata
    Exit Sub
    End If
    End If
    Next j
End If
End Sub

Private Sub gettag1()
Dim str As String
Dim str1 As String
Dim dataBuffer(256) As Byte
            Dim buf(256) As Byte
            Dim recbuff(256) As Byte
            Dim keyBuffer(1024) As Byte
            Dim PassWord(10) As Byte
            Dim Data(1024) As Byte
            Dim j As Integer, i As Byte
            Dim ret As Byte
            Dim CmdData(255) As Byte
            Dim btMemBank As Byte
            Dim btWordAdd As Byte
            Dim btWordCnt As Byte
            
            Dim TagCount(2) As Byte
            Dim DataLen(2) As Byte
            Dim ReadLen(2) As Byte
            Dim AntID(2) As Byte
            Dim ReadCount(2) As Byte
            
            Dim bytesData(255) As Byte
            Dim bytePCData(1) As Byte
            Dim byteEPCData(19) As Byte
            Dim byteReadData(19) As Byte
            Dim byteTIDData(1) As Byte
            Dim byteCRCData(1) As Byte
            Dim DataLenth As Long
            Dim ReadLenth As Long
            Dim EPCLenth As Long
            Dim DisData(255) As Byte
            
            ret = rdy_read(3, 0, 6, 0, TagCount(0), DataLen(0), Data(0), ReadLen(0), AntID(0), ReadCount(0))
            If ret <> 0 Then
                Texttag.Text = ""
                Exit Sub
            End If
            
            For j = 2 To 13
                str = str + Right$("00" & Hex$(Data(j)), 2)
            Next
If str <> Label1.Caption Then
    Texttag.Text = str
End If
End Sub

Sub checksock1()
If Winsock1.State <> "7" Then
Winsock1.Close
Winsock1.Connect
End If
End Sub

Sub checksock2()
If Winsock2.State <> "7" Then
Winsock2.Close
Winsock2.Connect
End If
End Sub

Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and  d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
recfound = True
End If
End Sub

Private Sub cmddel_Click()
cmdsave.Enabled = False
tctr = 0
If doorin = 1 Then
        openclose1
        d2 = 1
        door2.FillColor = &HFF00&
        If d1 = 1 Then
            openclose
            d1 = 0
            door1.FillColor = &HFF&
        End If
Else
        openclose
        d1 = 1
        door1.FillColor = &HFF00&
        If d2 = 1 Then
            openclose1
            d2 = 0
            door2.FillColor = &HFF&
        End If
End If
cmddel.Enabled = False
End Sub

Private Sub cmdref_Click()
unlocktxt
Text4.SetFocus
End Sub

Private Sub cmdsave_Click()
If CDbl(Text7.Text) < CDbl(Text9.Text) Then
    MsgBox "Weight cannot be more than RLW"
    Exit Sub
End If
a = MsgBox("Want to Update Records", vbOKCancel, "Update ?")
If a = 1 Then
    If recfound = False Then
        rs1.AddNew
    End If
    
    rs1.Fields("season").Value = Trim(Combo1.Text)
    rs1.Fields("d_serial").Value = Trim(Text1.Text)
    rs1.Fields("date_in").Value = CDate(DTPicker1.Value)
    rs1.Fields("time_in").Value = Trim(Text2.Text)
    rs1.Fields("o_name").Value = Trim(Text3.Text)
    rs1.Fields("V_no").Value = Trim(Text4.Text)
    rs1.Fields("c_name").Value = Trim(Text5.Text)
    rs1.Fields("c_address").Value = Trim(Text6.Text)
    rs1.Fields("RLW").Value = Trim(Text7.Text)
    rs1.Fields("material").Value = Trim(Text8.Text)
    rs1.Fields("first_wt").Value = Val(Text9.Text)
    rs1.Fields("second_wt").Value = 0
    rs1.Fields("tc_code").Value = Trim(Text11.Text)
    rs1.Fields("tm_code").Value = Trim(Text10.Text)
    rs1.Fields("tag").Value = Trim(Text12.Text)
    rs1.Update

        tctr = 0

If doorin = 1 Then
        openclose1
        d2 = 1
        door2.FillColor = &HFF00&
Else
        openclose
        d1 = 1
        door1.FillColor = &HFF00&
End If

    b = 1
    
    While b = 1
    b = MsgBox("Print Slip", vbOKCancel, "Print ?")
    If b = 1 Then
        'salebill_print
        'dosprint
        printrep
    End If
    Wend
End If
unlocktxt
Text4.SetFocus
End Sub

Private Sub salebill_print()
'setinfo
lctr = 0
Set wstream = fsys.OpenTextFile(App.Path & "\" & "rep.txt", ForWriting, True)
billprint
'wstream.Close
End Sub

Private Sub dosprint()
Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "\" & "r.bat", ForWriting, True)
wstream.WriteLine "type " + App.Path + "\rep.txt  >  prn"
wstream.Close
Shell App.Path + "\" + "r.bat"
End Sub


Private Sub printrep()
'On Error Resume Next
sifstwt.Sections("Section1").Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
sifstwt.Sections("Section1").Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
sifstwt.Sections("Section1").Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)

sifstwt.Sections("Section1").Controls("Label4").Caption = PADC("Simple First Weight Slip", 25)

Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set sifstwt.DataSource = rs1
sifstwt.Sections("Section1").Controls("Text1").DataField = "SEASON"
sifstwt.Sections("Section1").Controls("Label5").Caption = padl("Year   ", 20) & ": " & padl(rs1.Fields("SEASON").Value, 15)
sifstwt.Sections("Section1").Controls("Label6").Caption = padl("Date   ", 20) & ": " & padl(rs1.Fields("date_in").Value, 15)
sifstwt.Sections("Section1").Controls("Label7").Caption = padl("Time In", 20) & ": " & padl(rs1.Fields("Time_in").Value, 15)
sifstwt.Sections("Section1").Controls("Label8").Caption = PADR("Daily Serial", 20) & ": " & padl(rs1.Fields("d_serial").Value, 15)
sifstwt.Sections("Section1").Controls("Label9").Caption = PADR("Operator Name", 20) & ": " & padl(rs1.Fields("O_name").Value, 35)
sifstwt.Sections("Section1").Controls("Label10").Caption = PADR("Vehicle No", 20) & ": " & padl(rs1.Fields("v_no").Value, 35)
sifstwt.Sections("Section1").Controls("Label11").Caption = PADR("Name", 20) & ": " & padl(rs1.Fields("c_name").Value, 35)
sifstwt.Sections("Section1").Controls("Label12").Caption = PADR("Address", 20) & ": " & padl(rs1.Fields("c_address").Value, 35)
sifstwt.Sections("Section1").Controls("Label13").Caption = PADR("Colly/Mat Name", 20) & ": " & padl(rs1.Fields("Material").Value, 35)
sifstwt.Sections("Section1").Controls("Label14").Caption = PADR("First Weight: ", 13) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15)
sifstwt.Sections("Section1").Controls("Label15").Caption = PADR(COMPAUTH, 24) & padl("Weighing Operator", 20)
sifstwt.Sections("Section1").Controls("Label16").Caption = padl(compdes, 100)
Else
    sifstwt.Sections("Section1").Controls("Label4").Caption = PADC(Trim("No Such Serial No Created"), 30)
End If

sifstwt.Show vbModal

End Sub



Private Sub billprint()
'Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & Chr(14) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & Chr(14) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & Chr(18) & Chr(14) & PADC("Simple First Weight Slip", 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & String(48, "-")
 
Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    
    wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Year   ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("SEASON").Value, 15)
    wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date   ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15)
    wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15)
    
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("d_serial").Value, 15)
    
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 35)
    
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 35)
    
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_name").Value, 35)
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Address", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_address").Value, 35)
    
'    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Challan/S.O.No.", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("challan_so").Value, 35)
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colly/Mat Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Material").Value, 35)
    
    wstream.WriteBlankLines 1
    wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 13) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15) & Chr(27) & "F"
    wstream.WriteBlankLines 1
    wstream.WriteLine Chr(18) & String(48, "-")
    wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR(COMPAUTH, 24) & Chr(27) & "F" & padl("Weighing Operator", 20)
    wstream.WriteLine Chr(18) & String(48, "-")
    wstream.WriteLine Chr(15) & padl(compdes, 100) + Chr(27) & "F" & Chr(18)
Else
    wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim("No Such Serial No Created"), 30) & Chr(27) & "F"
End If

wstream.WriteBlankLines 3
wstream.Close
Set wstream = Nothing

End Sub


Private Sub sesiongen()
Combo1.Clear
Combo1.Refresh
If DTPicker1.Value <= CDate(Format(CDate("31/03/" + str(Year(DTPicker1.Value))), "mm/dd/yyyy")) Then
Combo1.AddItem Trim(Mid(Year(DTPicker1.Value), 1, 2)) + Right(str(Year(DTPicker1.Value) - 1), 2) + "-" + Trim(Mid(Year(DTPicker1.Value), 1, 2)) + Right(str(Year(DTPicker1.Value)), 2)
Else
Combo1.AddItem Trim(Mid(Year(DTPicker1.Value), 1, 2)) + Right(str(Year(DTPicker1.Value)), 2) + "-" + Trim(Mid(Year(DTPicker1.Value), 1, 2)) + Right(str(Year(DTPicker1.Value) + 1), 2)
End If
Combo1.ListIndex = 0
End Sub


Private Sub cmdsave_GotFocus()

If Trim(Text1.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text1.SetFocus
Exit Sub
End If


If Trim(Text2.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text2.SetFocus
Exit Sub
End If


If Trim(Text3.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text3.SetFocus
Exit Sub
End If

If Trim(Text4.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text4.SetFocus
Exit Sub
End If

If Trim(Text5.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text5.SetFocus
Exit Sub
End If


If Trim(Text6.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text6.SetFocus
Exit Sub
End If

If Trim(Text7.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text7.SetFocus
Exit Sub
End If



If Trim(Text8.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text8.SetFocus
Exit Sub
End If


If Val(Text9.Text) = 0 Then
MsgBox "Zero Value Can not be accepted", vbInformation, "Zero Weight Error"
Command1.SetFocus
Exit Sub
End If



If Trim(Text10.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text10.SetFocus
Exit Sub
End If


If Trim(Text11.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text11.SetFocus
Exit Sub
End If



If itmfound = False Then
Text10.SetFocus
MsgBox "Item Not found in master", vbCritical, "Item Error"
Exit Sub
End If


End Sub

Private Sub Command1_Click()

' checkempty1 As Boolean
' checkempty2  As Boolean
' chkwt1 As Double
' chkwt2 As Double
' chkunit As String

DATAREAD
'Text9.SetFocus
End Sub

Private Sub DATAREAD()
If checkempty1 = True Then
    If platformin = True Then
        If Val(main.comwt.Caption) > 0 Then
            If Val(Text9.Text) = 0 Then
                Text9.Text = main.comwt.Caption
            End If

            platformin = False
        Else
            MsgBox "There is No Weight", vbCritical, "Weight Error"
        End If
    Else
        MsgBox "Platform is Empty or not ready", vbCritical, "Error"
    End If
Else
                Text9.Text = main.comwt.Caption
End If
End Sub

Private Sub Command1_GotFocus()
List2.Visible = False
If Trim(Text11.Text) = "" Then
Text11.SetFocus
Exit Sub
End If

If Trim(Text10.Text) = "" Then
Text10.SetFocus
Exit Sub
End If


'If itmfound = False Then
'Text10.SetFocus
'Exit Sub
'End If
End Sub



Private Sub Command4_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
autoincr
sesiongen
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub Form_Activate()
If Val(main.comwt.Caption) > 20 Then
    MsgBox "Platform not empty or not ready"
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
username = loginname
Call Conn
unlocktxt
COMPSAT = "C.I.S.F."
itmhlp
achlp

doorin = 0
d1 = 0
d2 = 0
Call Conn1

If boomcode = 1 Then
    Winsock1.RemotePort = paths(5)
    Winsock1.RemoteHost = ips(5)
    Winsock1.Close
    Winsock1.Connect
    
    Winsock2.RemotePort = paths(7)
    Winsock2.RemoteHost = ips(7)
    Winsock2.Close
    Winsock2.Connect
End If
End Sub
Private Sub itmdatashow()
itmfound = False
Set rs = New ADODB.Recordset
rs.Open "Select * from mater1 where m_code='" & Trim(Text10.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    rs.MoveFirst
    itmfound = True
    Text10.Text = rs.Fields("m_code").Value
    Text8.Text = rs.Fields("m_name").Value
Else
    Text8.Text = ""
    MsgBox "Item Not find ", vbInformation, "Item Not Exist"
End If

End Sub

Private Sub custdata()
If Trim(Text11.Text) <> "" Then
    Set rs2 = New ADODB.Recordset
    rs2.Open "Select * from cust1 where c_code='" & Trim(Text11.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs2.RecordCount > 0 Then
        Text11.Text = rs2.Fields("C_code").Value
        Text5.Text = rs2.Fields("c_name").Value & ""
        Text6.Text = rs2.Fields("c_address").Value & ""
        Text10.SetFocus
    Else
        Text5.Text = ""
        Text6.Text = ""
        MsgBox "This code does not exist", vbInformation, "Code Not Found"
        Text11.SetFocus
    End If
Else
    Text5.Text = ""
    Text6.Text = ""
    Text11.Text = ""
End If

End Sub

Private Sub itmhlp()
Set rs = New ADODB.Recordset
rs.Open "Select * from mater1 order by m_name", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If
g = 0
ReDim tmcode(rs.RecordCount)
While Not rs.EOF
List2.AddItem rs.Fields("m_name").Value
tmcode(g) = rs.Fields("m_code").Value
g = g + 1
rs.MoveNext
Wend
End Sub



Private Sub achlp()
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cust1 order by c_name", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
End If
g = 0
ReDim tmpcode(rs2.RecordCount)
ReDim tmpcode1(rs2.RecordCount)
ReDim tmpcode2(rs2.RecordCount)
ReDim tmpcode3(rs2.RecordCount)
While Not rs2.EOF
List1.AddItem rs2.Fields("c_name").Value
tmpcode(g) = rs2.Fields("c_code").Value
tmpcode1(g) = rs2.Fields("c_address").Value
tmpcode2(g) = rs2.Fields("c_orderno").Value
tmpcode3(g) = rs2.Fields("o_date").Value
g = g + 1
rs2.MoveNext
Wend

End Sub

Private Sub autoincr()
Dim pribno As Integer
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# order by val(d_serial)", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveLast
pribno = Val(Right(rs1.Fields("d_serial").Value, 3))
pribno = pribno + 1
Text1.Text = IncrBillNo1(rs1.Fields("d_serial").Value)
Else
pribno = 1
End If
Text1.Text = Format(pribno, "00#")
End Sub

Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = Format(Time, "HH:MM")
Text3.Text = username
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text11.Text = ""
Text10.Text = ""
Text9.Text = Val(0)
List1.Visible = False
List2.Visible = False
DTPicker1.Value = Date
List1.Visible = False
List2.Visible = False
autoincr
sesiongen
End Sub

Private Sub List1_DblClick()
        Text5.Text = List1.List(List1.ListIndex)
        Text6.Text = tmpcode1(List1.ListIndex)
        'Text7.Text = tmpcode2(List1.ListIndex)
        Text11.Text = tmpcode(List1.ListIndex)
        pressf1 = False
        Text10.SetFocus
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call List1_DblClick
End If
End Sub

Private Sub List2_DblClick()
Text8.Text = List2.List(List2.ListIndex)
Text10.Text = tmcode(List2.ListIndex)
itmfound = True
pressf1 = False
Command1.SetFocus
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call List2_DblClick
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If


End Sub

Public Sub dd1()
If Len(str1) > 1 Then
    For X = 1 To Len(str1) + 1
        m = Mid(str1, X, X + 1)
        If m = " " Then
            num = 0
        Else
            num = num + 1
        End If
    Next X
End If

If Len(str1) = 0 Then
    num = 0
End If
End Sub
Public Sub KeyPress1(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then
If Len(str1) > 1 Then
              If num = 1 Then
                  g = KeyAscii - 32
                  KeyAscii = g
            Else
            g = KeyAscii
            KeyAscii = g
            End If
            
End If


If Len(str1) = 0 Then
  g = KeyAscii - 32
  KeyAscii = g
  End If
End If
End Sub

Private Sub Text10_Change()
'str1 = Text10.Text
'Call dd1
'If Trim(Text10.Text) <> "" Then
'itmdatashow
'End If
'
'For i = 0 To List2.ListCount - 1
'      If Trim(Text10.Text) = Left(List2.List(i), Len(Trim(Text10.Text))) Then
'                    List2.ListIndex = i
'                    Exit Sub
'        End If
'    Next i
End Sub

Private Sub Text10_GotFocus()
str1 = ""
num = 0
If Trim(Text11.Text) = "" Then
Text11.SetFocus
Exit Sub
End If

List1.Visible = False
For i = 0 To List2.ListCount - 1
      If Trim(Text10.Text) = Left(List2.List(i), Len(Trim(Text10.Text))) Then
                    List2.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text10.Text) <> "" Then
        If pressf1 = True Then
            Text8.Text = List2.List(List2.ListIndex)
            Text10.Text = tmcode(List2.ListIndex)
            pressf1 = False
            Command1.SetFocus
        Else
            itmdatashow
            Command1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
If List2.Visible = True Then
List2.Visible = False
pressf1 = False
Else
pressf1 = True
List2.Visible = True
End If
End If

If KeyCode = 40 Then
List2.SetFocus
End If



End Sub

Private Sub Text11_Change()
'str1 = Text11.Text
'Call dd1
'If pressf1 = False Then
' custdata
' Else
'For i = 0 To List1.ListCount - 1
'      If Trim(Text11.Text) = Left(List1.List(i), Len(Trim(Text11.Text))) Then
'                    List1.ListIndex = i
'                    Exit Sub
'        End If
'    Next i
'    End If
End Sub

Private Sub Text11_GotFocus()
str1 = ""
num = 0
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

For i = 0 To List1.ListCount - 1
      If Trim(Text11.Text) = Left(List1.List(i), Len(Trim(Text11.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text11.Text) <> "" Then
            If pressf1 = True Then
                Text5.Text = List1.List(List1.ListIndex)
'                Text6.Text = tmpcode1(List1.ListIndex)
                Text11.Text = tmpcode(List1.ListIndex)
                pressf1 = False
             Else
                custdata
                pressf1 = False
             End If
    End If
Else
    Call KeyPress1(KeyAscii)
End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
If List1.Visible = True Then
List1.Visible = False
pressf1 = False
Else
List1.Visible = True
pressf1 = True
End If
End If

If KeyCode = 40 Then
List1.SetFocus
End If


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Text4.Text) <> "" Then
Text7.SetFocus
End If
End If
End Sub

Private Sub Text5_Change()
str1 = Text5.Text
Call dd1
End Sub

Private Sub Text5_GotFocus()
str1 = ""
num = 0

Text5.SelStart = 0
Text5.SelLength = Len(Text1.Text)

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
Else
Call KeyPress1(KeyAscii)

End If
End Sub

Private Sub Text6_Change()
str1 = Text6.Text
Call dd1
End Sub

Private Sub Text6_GotFocus()
str1 = ""
num = 0
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.SetFocus
Else
Call KeyPress1(KeyAscii)
End If
End Sub


Private Sub Text7_GotFocus()
If Trim(Text4.Text) = "" Then
    Text4.SetFocus
    Exit Sub
End If

List1.Visible = False
List2.Visible = False
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)

End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Text7.Text) <> "" Then
Text11.SetFocus
End If
End If
End Sub

Private Sub Text8_Change()
str1 = Text8.Text
Call dd1
End Sub

Private Sub Text8_GotFocus()
str1 = ""
num = 0
Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
Else
Call KeyPress1(KeyAscii)

End If

End Sub

Private Sub Text9_GotFocus()
List1.Visible = False
List2.Visible = False
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

'If Val(Text9.Text) = 0 Then
'Command1.SetFocus
'Exit Sub
'End If


End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text9.Text) > 0 Then
cmdsave.SetFocus
Else
Command1.SetFocus
End If
End If

End Sub


Private Sub Timer1_Timer()
If Winsock1.State <> "7" And boomcode = 1 Then
Winsock1.Close
Winsock1.Connect
End If

If Winsock2.State <> "7" And boomcode = 1 Then
Winsock2.Close
Winsock2.Connect
End If

If main.comwt.Caption <= -20 Then
    MsgBox "Dont try again"
    Command1.Enabled = False
End If
Text2.Text = Format(Time, "HH:MM")

If wtmode = "D" Then
If main.comwt.Caption > 5000 Then
If doorin = 1 Then
    If d1 = 1 Then
        openclose
        d1 = 0
        door1.FillColor = &HFF&
    End If
Else
    If d2 = 1 Then
        openclose1
        d2 = 0
        door2.FillColor = &HFF&
    End If
End If
End If
End If

If wtmode = "R" Then
If main.comwt.Caption > 14000 Then
If doorin = 1 Then
    If d1 = 1 Then
        openclose
        d1 = 0
        door1.FillColor = &HFF&
    End If
Else
    If d2 = 1 Then
        openclose1
        d2 = 0
        door2.FillColor = &HFF&
    End If
End If
End If
End If


If main.comwt.Caption < 500 Then
If doorin = 1 Then
    If d2 = 1 Then
        Timer3.Enabled = True
        If tctr = 15 Then
            openclose1
            d2 = 0
            door2.FillColor = &HFF&
            tctr = 0
            Me.Timer3.Enabled = False
            Timer4.Enabled = True
        End If
    End If
ElseIf doorin = 2 Then
    If d1 = 1 Then
        Timer3.Enabled = True
        If tctr = 15 Then
            openclose
            d1 = 0
            door1.FillColor = &HFF&
            tctr = 0
            Me.Timer3.Enabled = False
            Timer4.Enabled = True
        End If
    End If
End If
End If
doorin1 = doorin
d11 = d1
d21 = d2
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If tagvalid = False Then
a = Net_Connect(unames(4), CLng(passws(4)), ips(4), CLng(paths(4)))
doorin = 1
getTag
End If

If tagvalid = False Then
a = Net_Connect(unames(6), CLng(passws(6)), ips(6), CLng(paths(6)))
doorin = 2
getTagg
End If
End Sub

Private Sub Timer3_Timer()
tctr = tctr + 1
Label29.Caption = "GATE CLOSING IN : " + CStr(15 - tctr)
End Sub

Private Sub Timer4_Timer()
tctr = tctr + 1
If tctr > 3 Then
Unload Me
End If
End Sub
