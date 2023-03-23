VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form sscwt 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   480
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   7200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   7200
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   720
      Top             =   7200
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   7200
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2400
      TabIndex        =   35
      Top             =   6240
      Width           =   8295
      Begin VB.CommandButton Command5 
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
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
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
      Left            =   10260
      TabIndex        =   5
      ToolTipText     =   "Unload Form"
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   8295
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox Text3 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   120
         Width           =   5415
      End
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   4320
         Width           =   1575
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   0
         Top             =   1440
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73793537
         CurrentDate     =   39961
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "TAG NUM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   2400
         TabIndex        =   43
         Top             =   5040
         Width           =   5775
      End
      Begin VB.Label Label31 
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
         Left            =   7680
         TabIndex        =   42
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label33 
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
         Left            =   6960
         TabIndex        =   41
         Top             =   4200
         Width           =   375
      End
      Begin VB.Shape door2 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7680
         Top             =   4500
         Width           =   375
      End
      Begin VB.Shape door1 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   6960
         Top             =   4500
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Out Time"
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
         Left            =   4440
         TabIndex        =   39
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label outtime 
         Caption         =   "Outtime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   38
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label18 
         Caption         =   "Out Date"
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
         Left            =   4440
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label odate 
         Caption         =   "Odate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   36
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label intime 
         Caption         =   "intime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   34
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label16 
         Caption         =   "In Time"
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
         Left            =   4440
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label sesson 
         Caption         =   "sesson"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label15 
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
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label ntw 
         Caption         =   "ntw"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   30
         Top             =   4800
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Net Weight"
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
         Left            =   720
         TabIndex        =   29
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Second Weight"
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
         Left            =   720
         TabIndex        =   27
         Top             =   4350
         Width           =   1455
      End
      Begin VB.Label op2 
         Caption         =   "op2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   26
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label op1 
         Caption         =   "op1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6480
         TabIndex        =   25
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Second  Operator"
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
         Left            =   4440
         TabIndex        =   24
         Top             =   3600
         Width           =   1845
      End
      Begin VB.Label Label9 
         Caption         =   "First Operator"
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
         TabIndex        =   23
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label fwt 
         Caption         =   "fwt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   22
         Top             =   3915
         Width           =   1185
      End
      Begin VB.Label material 
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   21
         Top             =   3480
         Width           =   3405
      End
      Begin VB.Label chno 
         Caption         =   "chno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   20
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label cadr 
         Caption         =   "cadr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   19
         Top             =   2520
         Width           =   3120
      End
      Begin VB.Label cname 
         Caption         =   "cname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   18
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label vno 
         Caption         =   "vno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   17
         Top             =   1815
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "First Weight"
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
         Left            =   720
         TabIndex        =   16
         Top             =   3885
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Material"
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
         Left            =   720
         TabIndex        =   15
         Top             =   3435
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "RLW"
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
         Left            =   720
         TabIndex        =   14
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
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
         Left            =   720
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Vehicle No"
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
         Left            =   720
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Serial No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "In Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Simple Second  Weight"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label13 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   8325
   End
End
Attribute VB_Name = "sscwt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim doorin As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim tctr As Integer
Dim wtmode As String
Dim tagtrips As Integer
Dim tripsdone As Integer

Public OutBufferCS, nBytesWrite As Byte
Dim TXBuffer(200) As Byte
Dim RXBuffer(100) As Byte
Dim FunctionValue As Byte
Dim waitAnswerTime As Integer
Dim lastCmd As Byte
Dim tagvalid As Boolean

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
On Error Resume Next
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
    'Addlog "Send:  ", s
End Sub


Private Sub DoSendData1()
On Error Resume Next
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
    'Addlog "Send:  ", s
End Sub

Private Sub Addlog(ByVal titl As String, ByVal msg As String)
    'Text3.Text = Text3.Text + (Chr(13)) + Chr(10) + titl + msg
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


'Function trimslash(tda As String) As String
'trimslash = Mid(tda, 1, 2) + Mid(tda, 4, 2) + Mid(tda, 7, 4)
'End Function

Function trimdate(tda As Date) As String
trimdate = Format(Day(tda), "0#") & Format(Month(tda), "0#") & Format(Year(tda), "0###")
End Function

Function DirExists(ByVal DirName As String) As Boolean
   On Error GoTo EH
   If (GetAttr(DirName) And vbDirectory) = vbDirectory Then
       DirExists = True
   End If
   Exit Function
EH:
   DirExists = False
End Function

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
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
End If

End Sub
Private Sub Command1_Click()
If CDbl(Text2.Text) > CDbl(chno.Caption) Then
    MsgBox "Weight cannot be more than RLW"
    Exit Sub
End If
a = MsgBox("Want to Update data", vbInformation, "Update   ?")
If a = 1 Then
dataset
rs1.Fields("o2_name").Value = op2.Caption
rs1.Fields("time_out").Value = outtime.Caption
rs1.Fields("date_out").Value = CDate(odate.Caption)
rs1.Fields("second_wt").Value = Val(Text2.Text)
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
        printrep
    End If
    Wend

unlocktxt
Text1.Text = ""
Text1.SetFocus
End If

End Sub

Private Sub printrep()
'On Error Resume Next
spsndwt.Sections("Section1").Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
spsndwt.Sections("Section1").Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
spsndwt.Sections("Section1").Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)

spsndwt.Sections("Section1").Controls("Label4").Caption = PADC("Simple First Weight Slip", 25)

Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set spsndwt.DataSource = rs1
spsndwt.Sections("Section1").Controls("Text1").DataField = "SEASON"
spsndwt.Sections("Section1").Controls("Label5").Caption = padl("Year   ", 15) & ": " & padl(rs1.Fields("SEASON").Value, 15)
spsndwt.Sections("Section1").Controls("Label6").Caption = padl("Date In", 15) & ": " & padl(rs1.Fields("date_in").Value, 15) & padl("Date Out", 10) & ": " & odate.Caption
spsndwt.Sections("Section1").Controls("Label7").Caption = padl("Time In", 15) & ": " & padl(rs1.Fields("Time_in").Value, 15) & padl("Time Out", 10) & ": " & outtime.Caption
spsndwt.Sections("Section1").Controls("Label8").Caption = PADR("Daily Serial", 15) & ": " & padl(rs1.Fields("d_serial").Value, 15)
spsndwt.Sections("Section1").Controls("Label9").Caption = PADR("Operator1", 15) & ": " & padl(rs1.Fields("O_name").Value, 20) & padl("Operator2", 10) & ": " & op2.Caption
spsndwt.Sections("Section1").Controls("Label10").Caption = PADR("Vehicle No", 15) & ": " & padl(rs1.Fields("v_no").Value, 35)
spsndwt.Sections("Section1").Controls("Label11").Caption = PADR("Name", 15) & ": " & padl(rs1.Fields("c_name").Value, 35)
spsndwt.Sections("Section1").Controls("Label12").Caption = PADR("Address", 15) & ": " & padl(rs1.Fields("c_address").Value, 35)
spsndwt.Sections("Section1").Controls("Label13").Caption = PADR("Colly/Mat Name", 15) & ": " & padl(rs1.Fields("Material").Value, 35)
spsndwt.Sections("Section1").Controls("Label14").Caption = PADR("First Weight: ", 13) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 12) & padl("Second Wt", 10) & padl(Text2.Text & " Kg", 15)
spsndwt.Sections("Section1").Controls("Label17").Caption = PADR("Net weight  : ", 13) & padl(ntw.Caption & " Kg", 15)
spsndwt.Sections("Section1").Controls("Label15").Caption = PADR(COMPAUTH, 44) & padl("Weighing Operator", 20)
spsndwt.Sections("Section1").Controls("Label16").Caption = padl(compdes, 100)
Else
    spsndwt.Sections("Section1").Controls("Label4").Caption = PADC(Trim("No Such Serial No Created"), 30)
End If

spsndwt.Show vbModal

End Sub


Private Sub printdata()
Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADC("Simple Second Weight Slip", 25) & Chr(27) & "F"
 wstream.WriteLine String(53, "-")
Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from smdtwtrep where date_out=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst

'wstream.WriteLine Chr(18) & "E" & Chr(14) & padl("Year   ", 10) & Chr(27) & "F" & Chr(18) & "E" & padl(rs1.Fields("SEASON").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("d_serial").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("1st Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("2nd Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O2_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Address", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_address").Value, 35)
'wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Challan/S.O.No.", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("challan_so").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colly/Mat Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("M_name").Value, 35)

wstream.WriteBlankLines 1

wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Date Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Time Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)



wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 14) & padl(":", 2) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Second Weight", 14) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Net Weight   ", 14) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine String(53, "-")
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR(COMPAUTH, 22) & Chr(27) & "F" & padl("Weighing Operator", 20)
wstream.WriteLine String(53, "-")
wstream.WriteLine Chr(15) & padl(compdes, 100)
Else
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim("No Such Serial No Created"), 30) & Chr(27) & "F"
End If
wstream.WriteBlankLines 3

 wstream.Close
 Set wstream = Nothing
End Sub

Private Sub dosprint()
Set fsys = CreateObject("scripting.filesystemobject")
Set wstream = fsys.OpenTextFile(App.Path & "\" & "r.bat", ForWriting, True)
wstream.WriteLine "type " + App.Path + "\" & "rep.txt  >  prn"
wstream.Close
Shell App.Path + "\" + "r.bat"
End Sub

Private Sub Command1_GotFocus()
If Val(Text2.Text) < 20 Then
MsgBox "Weight is too Low  ", vbInformation, "Weight is too low "
Command3.Enabled = True
Command3.SetFocus
Exit Sub
End If

If Trim(Text1.Text) = "" Then
MsgBox "Serial No Not Exist", vbCritical, "Error"
Text1.SetFocus
Exit Sub
Else
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Else
Text1.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub Command2_Click()
unlocktxt
End Sub

Private Sub Command3_Click()
If Val(main.comwt.Caption) > 0 Then
    Text2.Text = main.comwt.Caption
    Command3.Enabled = False
Else
    MsgBox "There is no Weight", vbCritical, "Weight Error"
End If
End Sub

Private Sub Command3_GotFocus()
If Trim(Text1.Text) = "" Then
Text1.SetFocus
Exit Sub
End If
'Text2.SetFocus
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Command1.Enabled = False
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
Command5.Enabled = False
End Sub

Private Sub Form_Activate()
If Val(main.comwt.Caption) > 20 Then
    MsgBox "Platform not empty or not ready"
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture

Conn
'username = "Rahul"
unlocktxt
DTPicker1.Value = Date
End Sub

Private Sub showtagdata()
    Timer2.Enabled = False
    
    Set rs4 = New ADODB.Recordset
    rs4.Open "Select * from tags where tagno = '" + Trim(Texttag) + "' and valid='2'", co1, adOpenKeyset, adLockOptimistic
    If rs4.RecordCount > 0 Then
        rs4.MoveFirst
        wtmode = rs4.Fields("mode").Value
    End If
    rs4.Close
    datashow
    Text18.SetFocus
End Sub

Private Sub unlocktxt()
username = loginname
op2.Caption = username
sesson.Caption = ""
vno.Caption = ""

cname.Caption = ""
cadr.Caption = ""
op1.Caption = ""
intime.Caption = ""
chno.Caption = ""
material.Caption = ""
fwt.Caption = ""
odate.Caption = Date
outtime.Caption = Format(Time, "hh:mm")
ntw.Caption = ""
Text2.Text = ""
End Sub


Private Sub datashow()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where tag='" & Trim(Texttag.Text) & "' and val(second_wt)=0 order by date_in desc", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
op2.Caption = username
sesson.Caption = rs1.Fields("season").Value
vno.Caption = rs1.Fields("V_no").Value
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cust1  where c_code='" & Trim(rs1.Fields("tc_code").Value) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
cname.Caption = rs2.Fields("c_name").Value
End If

cadr.Caption = rs1.Fields("c_address").Value
op1.Caption = rs1.Fields("o_name").Value
intime.Caption = rs1.Fields("time_in").Value
chno.Caption = rs1.Fields("RLW").Value

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from Mater1  where m_code='" & Trim(rs1.Fields("tm_code").Value) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst

material.Caption = rs2.Fields("m_name").Value
End If

fwt.Caption = rs1.Fields("first_wt").Value

If Val(rs1.Fields("SECOND_WT") & "") > 0 Then
MsgBox "Dont Try Again", vbCritical, "Second Weight Already Taken"
Text1.Text = ""
unlocktxt
Exit Sub
End If

Else
unlocktxt
MsgBox "TAG Not Found", vbInformation, "Wrong Tag Number"
End If
Exit Sub

errm:

End Sub


Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Text1.Text) <> "" Then
datashow
If Command3.Enabled = True Then Command3.SetFocus
End If
End If

End Sub

Private Sub Text2_Change()
ntw.Caption = Abs(Val(fwt.Caption) - Val(Text2.Text))
If CLng(ntw.Caption) < 100 And Val(Text2.Text) > 0 Then
    ntw.Caption = 0
    Text2.Text = fwt.Caption
    MsgBox "Net weight too low, 0 weight taken"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text2.Text) = 0 Then
    If Command3.Enabled = True Then Command3.SetFocus
Else
Command1.SetFocus
End If
End If

End Sub


Private Sub Timer1_Timer()
On Error Resume Next

If Winsock1.State <> "7" And boomcode = 1 Then
Winsock1.Close
Winsock1.Connect
End If

If Winsock2.State <> "7" And boomcode = 1 Then
Winsock2.Close
Winsock2.Connect
End If


If main.comwt.Caption <= -20 Then
    main.Label4.Caption = "Dont try again"
    Command1.Enabled = False
    Command6.Enabled = False
End If

If wtmode = "D" Then
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

If wtmode = "R" Then
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


If main.comwt.Caption < 500 Then
If doorin = 1 Then
    If d2 = 1 Then
        Timer3.Enabled = True
        If tctr = 15 Then
            openclose1
            d2 = 0
            door2.FillColor = &HFF&
            tctr = 0
            Timer3.Enabled = False
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
            Timer3.Enabled = False
            Timer4.Enabled = True
        End If
    End If
End If
End If

outtime.Caption = Format(Time, "hh:mm")
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
Label30.Caption = "GATE CLOSING IN : " + CStr(15 - tctr)
End Sub

Private Sub Timer4_Timer()
tctr = tctr + 1
If tctr > 3 Then
Unload Me
End If
End Sub
