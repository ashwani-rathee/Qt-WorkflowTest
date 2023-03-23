VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H00E0E0E0&
   Caption         =   "WEIGHTSMART Ver 12.08"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   1125
   ClientWidth     =   15120
   Icon            =   "main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "main.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock111 
      Left            =   840
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.1.252"
      RemotePort      =   80
   End
   Begin VB.Timer Timer4 
      Interval        =   3000
      Left            =   5280
      Top             =   2880
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   2880
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      Height          =   450
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   15060
      TabIndex        =   9
      Top             =   1395
      Width           =   15120
      Begin VB.ListBox lb 
         Height          =   255
         Left            =   14160
         TabIndex        =   25
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   42725
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   42725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   42725
      End
      Begin VB.ComboBox readtags 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   60
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox invalidtags 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   60
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   42725
      End
      Begin VB.Label tagg 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   13200
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TAG :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.Label texttagg 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   6720
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label texttagg1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6840
         TabIndex        =   16
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   4320
      Top             =   2880
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   2880
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H000000A0&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   15060
      TabIndex        =   5
      Top             =   7305
      Width           =   15120
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11400
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   14460
         TabIndex        =   8
         Top             =   60
         Width           =   105
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   13410
         TabIndex        =   7
         Top             =   60
         Width           =   765
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE ENSURE THAT GATES ARE CLOSED BEFORE EVERY WEIGHMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   30
         Width           =   13215
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   4560
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
      DataBits        =   7
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   5280
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   1200
      DataBits        =   7
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H000000A0&
      ForeColor       =   &H00000080&
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   15060
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin VB.Label comwt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1395
         Left            =   10680
         TabIndex        =   4
         Top             =   -140
         Width           =   3975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MOVING PENDING DATA TO HEADQUARTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   8160
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   5280
         TabIndex        =   3
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   5280
         TabIndex        =   2
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   5280
         TabIndex        =   1
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15120
      TabIndex        =   24
      Top             =   0
      Width           =   15120
   End
   Begin MSCommLib.MSComm MSComm4 
      Left            =   6000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
      DataBits        =   7
   End
   Begin VB.Menu mastmnu 
      Caption         =   "      CONFIGURATION      "
      Begin VB.Menu compinf 
         Caption         =   "COMPANY"
      End
      Begin VB.Menu LN3 
         Caption         =   "-"
      End
      Begin VB.Menu usmast 
         Caption         =   "USERS"
      End
      Begin VB.Menu ln0 
         Caption         =   "-"
      End
      Begin VB.Menu csw 
         Caption         =   "SOFTWARE CONFIGURATION"
      End
      Begin VB.Menu ln0011 
         Caption         =   "-"
      End
      Begin VB.Menu chmac 
         Caption         =   "MACHINE"
      End
   End
   Begin VB.Menu areamenu 
      Caption         =   "      AREA      "
      Begin VB.Menu tagissuem 
         Caption         =   "TAG ISSUE"
      End
      Begin VB.Menu tagln1 
         Caption         =   "-"
      End
      Begin VB.Menu tagchangem 
         Caption         =   "TAG / TRANSACTION CANCEL"
      End
      Begin VB.Menu tagln2 
         Caption         =   "-"
      End
      Begin VB.Menu tagreportm 
         Caption         =   "TAG RE ISSUE / MODIFY"
      End
      Begin VB.Menu tagln4 
         Caption         =   "-"
      End
      Begin VB.Menu tagchange 
         Caption         =   "TAG CHANGE"
      End
      Begin VB.Menu tgchsp 
         Caption         =   "-"
      End
      Begin VB.Menu tagisre 
         Caption         =   "TAG ISSUE REPORT"
      End
   End
   Begin VB.Menu passch 
      Caption         =   "      CHANGE PASSWORD      "
   End
   Begin VB.Menu cadata 
      Caption         =   "      CA DATA      "
      Begin VB.Menu caimport 
         Caption         =   "CA DATA IMPORT"
      End
      Begin VB.Menu camspace 
         Caption         =   "-"
      End
      Begin VB.Menu cdma 
         Caption         =   "CA DATA MANUAL"
      End
      Begin VB.Menu ln5 
         Caption         =   "-"
      End
   End
   Begin VB.Menu dodata 
      Caption         =   "      DO DATA      "
      Begin VB.Menu doalo 
         Caption         =   "DO ALLOTMENT"
      End
      Begin VB.Menu cdmasp 
         Caption         =   "-"
      End
      Begin VB.Menu doalim 
         Caption         =   "DO ALLOTMENT IMPORT"
      End
      Begin VB.Menu sp222 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mastab 
      Caption         =   "      MASTER TABLES      "
      Begin VB.Menu excdata 
         Caption         =   "UNIT EXCISE DATA"
      End
      Begin VB.Menu spexc 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu do 
         Caption         =   "DELIVERY ORDER"
         Visible         =   0   'False
      End
      Begin VB.Menu doasp 
         Caption         =   "-"
      End
      Begin VB.Menu dor 
         Caption         =   "DOR"
      End
      Begin VB.Menu spacer77 
         Caption         =   "-"
      End
      Begin VB.Menu imast 
         Caption         =   "ITEM MASTER"
      End
      Begin VB.Menu ln8 
         Caption         =   "-"
      End
      Begin VB.Menu custmast 
         Caption         =   "CUSTOMER"
      End
      Begin VB.Menu ln9 
         Caption         =   "-"
      End
      Begin VB.Menu collmas 
         Caption         =   "COLLIERY"
      End
   End
   Begin VB.Menu smmnu 
      Caption         =   "      SIMPLE      "
      Begin VB.Menu smfstwt 
         Caption         =   "FIRST WEIGHT"
      End
      Begin VB.Menu smsp1 
         Caption         =   "-"
      End
      Begin VB.Menu SmSwt 
         Caption         =   "SECOND WEIGHT"
      End
      Begin VB.Menu smsp2 
         Caption         =   "-"
      End
      Begin VB.Menu smpitmast 
         Caption         =   "ITEM MASTER"
      End
      Begin VB.Menu smsp3 
         Caption         =   "-"
      End
      Begin VB.Menu cm2 
         Caption         =   "CUSTOMER MASTER"
      End
      Begin VB.Menu smsp4 
         Caption         =   "-"
      End
      Begin VB.Menu svf 
         Caption         =   "SIMPLE VIEW FILE"
      End
      Begin VB.Menu SMSP6 
         Caption         =   "-"
      End
      Begin VB.Menu owsp 
         Caption         =   "OLD WEIGH SLIP"
      End
   End
   Begin VB.Menu specmnu 
      Caption         =   "      WEIGHMENT      "
      Begin VB.Menu spfwet 
         Caption         =   "FIRST WEIGHT"
         Visible         =   0   'False
      End
      Begin VB.Menu LN6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu spscwt 
         Caption         =   "SECONT WEIGHT"
         Visible         =   0   'False
      End
      Begin VB.Menu spln1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu spv 
         Caption         =   "VIEW FILE"
      End
      Begin VB.Menu samln2 
         Caption         =   "-"
      End
      Begin VB.Menu osp 
         Caption         =   "OLD SLIP"
      End
      Begin VB.Menu wesp1 
         Caption         =   "-"
      End
      Begin VB.Menu traveh 
         Caption         =   "TRACK VEHICLE"
      End
   End
   Begin VB.Menu dreps 
      Caption         =   "      REPORTS      "
      Begin VB.Menu cwr 
         Caption         =   "CUSTOMER WISE"
      End
      Begin VB.Menu ln12 
         Caption         =   "-"
      End
      Begin VB.Menu materwrepo 
         Caption         =   "MATERIAL WISE"
      End
      Begin VB.Menu space113 
         Caption         =   "-"
      End
      Begin VB.Menu dobrep 
         Caption         =   "DO BALANCE"
      End
      Begin VB.Menu space124 
         Caption         =   "-"
      End
      Begin VB.Menu cwrep 
         Caption         =   "COLLIERY WISE"
      End
      Begin VB.Menu spacer11 
         Caption         =   "-"
      End
      Begin VB.Menu testre 
         Caption         =   "TEST"
      End
      Begin VB.Menu ln22 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu expdat 
         Caption         =   "EXPORT DATA"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu qmnu 
      Caption         =   "      EXIT      "
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Integer
Dim doorinn As Integer
Dim unitc As String
Dim mctr As Integer
Dim tmptagg As String


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
    If Trim(str) <> texttagg.Caption Then
    If validtag = False Then
        Timer3.Enabled = False
        texttagg.Caption = str
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
    If Trim(str) <> texttagg.Caption Then
    If validtag = False Then
        Timer3.Enabled = False
        texttagg.Caption = str
    End If
    End If
    Next j
End If
End Sub

Private Sub caimport_Click()
unloadfrm
caimportt.Show
End Sub

Private Sub cdma_Click()
unloadfrm
camanual.Show
End Sub

Private Sub chmac_Click()
unloadfrm
selmac.Show
End Sub

Private Sub cm2_Click()
unloadfrm
achead1.Show
End Sub

Private Sub collmas_Click()
unloadfrm
ItemMast1.Show
End Sub


Private Sub Command2_Click()
If Timer3.Enabled = False Then
    Timer3.Enabled = True
    Command2.BackColor = &HFF00&
Else
    Timer3.Enabled = False
    Command2.BackColor = &HFF&
End If
End Sub

Private Sub compinf_Click()
unloadfrm
cinfo.Show
End Sub

Private Sub comwt_Change()
Dim fso As New FileSystemObject
Set wstream = fsys.OpenTextFile("cc.txt", ForWriting, True)
wstream.WriteLine comwt.Caption
wstream.Close
str1 = App.Path + "\cc.txt"
str2 = App.Path + "\cc.dbf"
fso.CopyFile str1, str2, True
End Sub

Private Sub csw_Click()
unloadfrm
sfeature.Show
End Sub

Private Sub custmast_Click()
unloadfrm
achead.Show
End Sub

Private Sub cwr_Click()
unloadfrm
dcustomrep.Show
End Sub

Private Sub cwrsmp_Click()
unloadfrm
smpcustomrep.Show

End Sub

Private Sub cwrep_Click()
unloadfrm
dcollrep.Show
End Sub

Private Sub do_Click()
unloadfrm
D_order.Show
End Sub

Private Sub dwrsm_Click()
unloadfrm
smpdatwrep.Show
End Sub

Private Sub doalim_Click()
unloadfrm
doimport.Show
End Sub

Private Sub doalo_Click()
unloadfrm
allotment.Show
End Sub

Private Sub dobrep_Click()
unloadfrm
ddobalance.Show
End Sub

Private Sub dor_Click()
unloadfrm
D_oreduce.Show
End Sub

Private Sub dwio_Click()
Shell "DHNBCCL.EXE", vbMaximizedFocus
End Sub

Private Sub excdata_Click()
unloadfrm
excisedata.Show
End Sub

Private Sub expdat_Click()
unloadfrm
Export.Show
End Sub

Private Sub imast_Click()
unloadfrm
ItemMast.Show
End Sub

Private Sub compinfo()
Call Conn
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from party ", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
unitc = Format(rs1.Fields("AREACODE"), "0#") + Format(rs1.Fields("WBCODE").Value, "0#")
Label1.Caption = Trim(rs1.Fields("pname").Value) & ""
Label2.Caption = Trim(rs1.Fields("pinf").Value) & " Weighbridge - " + rs1.Fields("WBCODE").Value & ""
Label3.Caption = Trim(rs1.Fields("Padd").Value) & " (" + rs1.Fields("AREACODE").Value & "" + ")"


'Label1.Caption = "TESTING ONLY"
'Label2.Caption = "TESTING ONLY"
'Label3.Caption = "TESTING ONLY"

arcode = Format(rs1.Fields("AREACODE").Value, "0#") & ""
wbcode = rs1.Fields("WBCODE").Value & ""
End If
End Sub

Private Sub Iwr_Click()
unloadfrm
spItmwrep.Show
End Sub

Private Sub iwrsmp_Click()
unloadfrm
smpItmwrep.Show
End Sub


Private Sub itwr_Click()
unloadfrm
spItmwrep.Show
End Sub

Private Sub materwrepo_Click()
unloadfrm
dmaterrep.Show
End Sub

Private Sub MDIForm_Load()
On Error Resume Next

compinfo
Dim a As String
Dim i As Integer

validtag = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from featur", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst

If Mid(rs1.Fields("permis").Value, 1, 1) = "1" Then
smmnu.Visible = True
dwrsm.Visible = True
ln13.Visible = True
iwrsmp.Visible = True
ln14.Visible = True
cwrsmp.Visible = True

Else
smmnu.Visible = False
dwrsm.Visible = False
ln13.Visible = False
iwrsmp.Visible = False
ln14.Visible = False
cwrsmp.Visible = False
End If

If Mid(rs1.Fields("permis").Value, 2, 1) = "1" Then
specmnu.Visible = True
'spdwreport.Visible = True
''LN10.Visible = True
'iwr.Visible = True
'ln11.Visible = True
cwr.Visible = True
chmac.Visible = True
ln12.Visible = True

Else
specmnu.Visible = False
spdwreport.Visible = False
LN10.Visible = False
iwr.Visible = False
ln11.Visible = False
cwr.Visible = False
'ln12.Visible = False
spln1.Visible = False
chmac.Visible = False
spcln2.Visible = False
End If

checkempty1 = True
checkempty2 = True

chkwt1 = Val(rs1.Fields("smpwtr").Value)
chkwt2 = Val(rs1.Fields("spcwtr").Value)
chkunit = rs1.Fields("munit").Value
compdes = rs1.Fields("des").Value
COMPAUTH = rs1.Fields("AUTH").Value
' checkempty1 As Boolean
' checkempty2  As Boolean
' chkwt1 As Double
' chkwt2 As Double
' chkunit As String
End If


If logintype = "Admin" Then
compinf.Visible = True
LN3.Visible = True
usmast.Visible = True
ln0.Visible = True
csw.Visible = True
chmac.Visible = True
passch.Visible = False
Timer3.Enabled = False
End If

If logintype = "Manager" Then
compinf.Visible = False
LN3.Visible = False
usmast.Visible = True
ln0.Visible = False
csw.Visible = False
chmac.Visible = False
Timer3.Enabled = False
End If

If UCase(loginname) = UCase("cadata") Or logintype = "Admin" Or logintype = "Manager" Then
    cadata.Visible = True
    Timer3.Enabled = False
Else
    cadata.Visible = False
End If

If UCase(Left(loginname, 8)) = UCase("areadata") Or logintype = "Admin" Then
    areamenu.Visible = True
    specmnu.Visible = False
    dreps.Visible = False
    mastab.Visible = False
    Timer3.Enabled = False
    main.comwt.Visible = False
Else
    areamenu.Visible = False
End If

If UCase(Left(loginname, 8)) = UCase("testdata") Then
    areamenu.Visible = False
    specmnu.Visible = False
    dreps.Visible = False
    mastab.Visible = False
    Timer3.Enabled = False
    main.comwt.Visible = True
End If

If UCase(loginname) = UCase("dodata") Or logintype = "Admin" Or logintype = "Manager" Then
    dodata.Visible = True
    Timer3.Enabled = False
Else
    dodata.Visible = False
End If

If UCase(loginname) = UCase("sidata") Then
'    smmnu.Visible = True
'    specmnu.Visible = False
'    dreps.Visible = False
'    mastab.Visible = False
'
'    Timer3.Enabled = False
Else
    smmnu.Visible = False
End If

Set rs1 = New ADODB.Recordset
rs1.Open "Select * from paths order by id", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
i = 0
While rs1.EOF = False
    If Trim(rs1.Fields(1)) <> "" Then
       ips(i) = rs1.Fields(1)
    End If
    If Trim(rs1.Fields(2)) <> "" Then
    paths(i) = rs1.Fields(2)
    End If

    If Trim(rs1.Fields(3)) <> "" Then
    unames(i) = rs1.Fields(3)
    End If

    If Trim(rs1.Fields(4)) <> "" Then
    passws(i) = rs1.Fields(4)
    End If
    i = i + 1
    rs1.MoveNext
Wend
End If

rs1.Close

mctr = 0

If logintype = "Operator" Then
mastmnu.Visible = False
End If

If Time > shifta1 And Time <= shifta2 Then
    Label6.Caption = "A"
ElseIf Time > shiftb1 And Time <= shiftb2 Then
    Label6.Caption = "B"
Else
    Label6.Caption = "C"
End If

machcode = "1"
Call Conn
Set rs2 = New ADODB.Recordset
rs2.Open "Select * from oper", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
rs2.MoveFirst
machcode = rs2.Fields("oname")

If machcode = "1" Or machcode = "2" Or machcode = "4" Or machcode = "6" Then
    MSComm1.CommPort = 1
    MSComm1.PortOpen = True
    Timer2.Enabled = False
    Timer1.Enabled = True
ElseIf machcode = "5" Then
    MSComm3.CommPort = 3
    MSComm3.PortOpen = True
    Timer1.Enabled = False
    Timer2.Enabled = True
ElseIf machcode = "8" Then
    MSComm4.CommPort = 1
    MSComm4.PortOpen = True
    Timer1.Enabled = True
    Timer2.Enabled = False
ElseIf machcode = "7" Then
    Winsock111.RemoteHost = ips(8)
    Winsock111.RemotePort = paths(8)
    Winsock111.Connect
Else
    MSComm2.CommPort = 7
    MSComm2.PortOpen = True
    Timer2.Enabled = False
    Timer1.Enabled = True
End If

If IsNumeric(rs2.Fields("uname")) Then
    camcode = CInt(rs2.Fields("uname"))
Else
    camcode = 0
End If

If IsNumeric(rs2.Fields("pword")) Then
    boomcode = CInt(Left(rs2.Fields("pword"), 1))
Else
    boomcode = 1
End If

If IsNumeric(rs2.Fields("pword")) Then
    rfidcode = CInt(Mid(rs2.Fields("pword"), 2, 1))
Else
    rfidcode = 1
End If

If IsNumeric(rs2.Fields("pword")) Then
    If CInt(rs2.Fields("pword")) > 99 Then
        gpscode = CInt(Mid(rs2.Fields("pword"), 3, 1))
    Else
        gpscode = 0
    End If
Else
    gpscode = 0
End If


End If
rs2.Close

If camcode = 1 Then
    cctv.Show
End If

Timer3.Enabled = False

datact = 0
Call Conn2
Set rs3 = New ADODB.Recordset
rs3.Open "Select * from specialtmp", co2, adOpenKeyset, adLockOptimistic
If rs3.RecordCount > 0 Then
datact = rs3.RecordCount
End If
rs3.Close

If logintype = "Operator" And UCase(loginname) <> UCase("sidata") And UCase(Left(loginname, 8)) <> UCase("areadata") And UCase(loginname) <> UCase("cadata") And UCase(loginname) <> UCase("dodata") Then
    If rfidcode = 1 Then
        Conn1
        If datact = 0 Then
            Timer3.Enabled = True
            traveh.Visible = True
        Else
            Timer3.Enabled = False
            traveh.Visible = False
            movedata
        End If
    Else
            spfwet.Visible = True
            LN6.Visible = True
            spscwt.Visible = True
            spln1.Visible = True
    End If
Else
    Picture3.Visible = False
    traveh.Visible = False
End If

If UCase(Left(loginname, 8)) = UCase("testdata") Then
    areamenu.Visible = False
    specmnu.Visible = False
    dreps.Visible = True
    mastab.Visible = False
    Timer3.Enabled = False
    main.comwt.Visible = True
End If


invalidtags.Clear

'KillApp ("conncheck.exe")
'Shell ("conncheck.exe")
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
KillApp ("conncheck.exe")
End
End Sub

Sub movedata()
    comwt.Visible = False
    Label7.Visible = True
    Label9.Visible = True
    Conn1
    Dim tripsdone As Integer
    Dim tagtrips As Integer

    Set rs3 = New ADODB.Recordset
    rs3.Open "Select * from specialtmp", co2, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        rs3.MoveLast
        rs3.MoveFirst
        Label9.Caption = rs3.RecordCount
        While rs3.EOF = False
            Label9.Caption = Val(Label9.Caption) - 1
'            DTPicker2.Value = rs3.Fields("date_in").Value
'            DTPicker3.Value = rs3.Fields("date_out").Value
            If Not IsNull(rs3.Fields("date_in").Value) Then
                DTPicker2.Value = rs3.Fields("date_in").Value
            End If

            If rs3.Fields("second_wt").Value > 0 Then
                DTPicker3.Value = rs3.Fields("date_out").Value
            End If
            If rs3.Fields("second_wt").Value > 0 Then
                DTPicker4.Value = CDate(rs3.Fields("challan_date").Value)
            End If

            Set rs4 = New ADODB.Recordset
            rs4.Open "Select * from special where sl_no='" + rs3.Fields("sl_no") + "'", co1, adOpenKeyset, adLockOptimistic
            On Error GoTo errm

            If rs4.RecordCount = 0 Then
                costr = "insert into special(season,sl_no,date_in,time_in,tc_code,v_no,o_name,tm_code,first_wt,second_wt,rlw,do_no,coll_code,order_qty,balance_qty,dest,shift_in,tag)" _
                + " values ('" + rs3.Fields("season") + "','" + rs3.Fields("sl_no") + "','" _
                + Format(Year(DTPicker2.Value), "00##") + Format(Month(DTPicker2.Value), "0#") + Format(Day(DTPicker2.Value), "0#") + "','" _
                + rs3.Fields("time_in") + "','" + rs3.Fields("tc_code") + "','" + rs3.Fields("v_no") + "','" _
                + rs3.Fields("o_name") + "','" + rs3.Fields("tm_code") + "'," + CStr(rs3.Fields("first_wt")) + "," + CStr(rs3.Fields("second_wt")) + "," + CStr(rs3.Fields("rlw")) + ",'" _
                + rs3.Fields("do_no") + "','" + rs3.Fields("coll_code") + "'," + CStr(rs3.Fields("order_qty")) _
                + "," + CStr(rs3.Fields("balance_qty")) + ",'" + rs3.Fields("dest") + "','" + rs3.Fields("shift_in") + "','XXX')"
                co1.Execute costr
            End If

            If rs3.Fields("second_wt").Value > 0 Then
                costr = "update special set o2_name='" + rs3.Fields("o2_name") + "', time_out='" + rs3.Fields("time_out") _
                + "', date_out='" + Format(Year(DTPicker3.Value), "00##") + Format(Month(DTPicker3.Value), "0#") + Format(Day(DTPicker3.Value), "0#") _
                + "',second_wt=" + CStr(rs3.Fields("second_wt")) + ", balance_qty=" + CStr(rs3.Fields("balance_qty")) _
                + ", challan_date='" + Format(Year(DTPicker4.Value), "00##") + Format(Month(DTPicker4.Value), "0#") + Format(Day(DTPicker4.Value), "0#") _
                + "', challan_no='" + challan_no + "' where sl_no='" + rs3.Fields("sl_no") + "'"
                co1.Execute costr

                If Len(rs3.Fields("tag").Value) > 0 Then
                    Set rs5 = New ADODB.Recordset
                    rs5.Open "Select * from tags where tagno = '" + Trim(Texttag) + "' and valid='2'", co1, adOpenKeyset, adLockOptimistic
                    If rs5.RecordCount > 0 Then
                        rs5.MoveFirst
                        tagtrips = Val(rs5.Fields("tagtrips"))
                        tripsdone = Val(rs5.Fields("trips_done")) + 1
                        If Not IsNumeric(tripsdone) Then
                            tripsdone = 1
                        End If
                    End If
                    rs5.Close

                    If tripsdone >= tagtrips Then
                        co1.Execute "update tags set valid='0',trips_done=" + CStr(tripsdone) + " where tagno = '" + Trim(Texttag.Text) + "' and tsno='" + Trim(Text1.Text) + "'"
                    Else
                        co1.Execute "update tags set valid='1',trips_done=" + CStr(tripsdone) + " where tagno = '" + Trim(Texttag.Text) + "' and tsno='" + Trim(Text1.Text) + "'"
                    End If
                End If
            End If

            rs3.Delete
            rs3.MoveNext
        Wend
        MsgBox "Data transferred successfully"
    End If
    rs3.Close
    Exit Sub

errm:
    MsgBox Err.Description
End Sub

Private Sub osp_Click()
unloadfrm
spcwtprn.Show
End Sub

Private Sub owsp_Click()
samwtprn.Show
End Sub

Private Sub passch_Click()
unloadfrm
passwd.Show
End Sub

Private Sub qmnu_Click()
a = MsgBox("Want to Exit Application", vbOKCancel, "Exit to Application")
If a = 1 Then
End
End If

End Sub

Private Sub smfstwt_Click()
unloadfrm
sfstwt.Show
End Sub

Private Sub smpitmast_Click()
unloadfrm
'ItemMast1.Show
ItemMastsmp.Show
End Sub

Private Sub SmSwt_Click()
unloadfrm
sscwt.Show
End Sub

Private Sub spcosp_Click()
unloadfrm
spcwtprn.Show
End Sub

Private Sub spdwreport_Click()
spdatwrep.Show
End Sub

Private Sub spfwet_Click()
unloadfrm
spfstwt.Show
End Sub

Private Sub spscwt_Click()
unloadfrm
spswt.Show
End Sub

Public Sub unloadfrm()
Unload spcwtprn
Unload samwtprn
Unload achead
Unload sfeature
Unload achead1
Unload cinfo
Unload ItemMast
Unload ItemMast1
Unload sfstwt
Unload spfstwt
Unload spswt
Unload sscwt
Unload Users
Unload dcustomrep
Unload spdatwrep
Unload spItmwrep

Unload smpcustomrep
Unload smpdatwrep
Unload smpItmwrep
End Sub

Private Sub spv_Click()
spacialView.Show
End Sub

Private Sub svf_Click()
SMVIEW.Show
End Sub

Private Sub tagchange_Click()
unloadfrm
tagrchange.Show
End Sub

Private Sub tagchangem_Click()
unloadfrm
tag_excep.Show
End Sub

Private Sub tagisre_Click()
unloadfrm
tagreport.Show
End Sub

Private Sub tagissuem_Click()
unloadfrm
tagissue.Show
End Sub

Private Sub tagreportm_Click()
unloadfrm
tagreissue.Show
End Sub

Private Sub testre_Click()
unloadfrm
Form1.Show
End Sub


Private Sub Texttagg_Change()
On Error Resume Next
tmptagg = texttagg.Caption
If texttagg.Caption = "" Then
Exit Sub
End If

Timer3.Enabled = False
texttagg1.Caption = Left(tmptagg, 4) + "XXXXXXXXXXXXXXXXXX" + Right(tmptagg, 4)

Set rs4 = New ADODB.Recordset
rs4.Open "Select * from tags where tagno = '" + Trim(tmptagg) + "' and (valid='1' or valid='2')", co1, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
rs4.MoveFirst
    DTPicker1.Value = CDate(rs4.Fields("expiry").Value)
    If DTPicker1.Value > Date Then
                tagg.Caption = "EXPIRED TAG"
                Timer3.Enabled = True
    End If
    If rs4.Fields("unit").Value = unitc Then
        'If rs4.Fields("wmode").Value = "2" Then
'            If rs4.Fields("valid").Value = "1" Then
'                validtag = True
'                tagg.Caption = "SI 1st WT"
'                unloadfrm
'                sfstwt.Show
'            ElseIf rs4.Fields("valid").Value = "2" Then
'                validtag = True
'                tagg.Caption = "SI 2nd WT"
'                unloadfrm
'                sscwt.Show
'            Else
'                tagg.Caption = "INVALID TAG"
'                Timer3.Enabled = True
'            End If
'        Else
            If rs4.Fields("valid").Value = "1" Then
                Timer3.Enabled = False
                validtag = True
                texttagg.Caption = tmptagg
                tagg.Caption = "SP 1st WT"
                unloadfrm
                spfstwt.Show
            ElseIf rs4.Fields("valid").Value = "2" Then
                Timer3.Enabled = False
                validtag = True
                texttagg.Caption = tmptagg
                tagg.Caption = "SP 2nd WT"
                unloadfrm
                spswt.Show
            Else
                tagg.Caption = "INVALID TAG"
                Timer3.Enabled = True
            End If
        'End If
    Else
        tagg.Caption = "WRONG WB"
        Timer3.Enabled = True
    End If
Else
    tagg.Caption = "INVALID TAG"
    invalidtags.AddItem Trim(texttagg.Caption)
    Timer3.Enabled = True
End If
rs4.Close
End Sub



Private Sub Timer1_Timer()
Dim i, b As Long
Dim a As String
On Error Resume Next

If Timer3.Enabled = True Then
    Command2.BackColor = &HFF00&
Else
    Command2.BackColor = &HFF&
End If

If machcode = "5" Then
    Exit Sub
End If
If machcode = "2" Then
    a = MSComm1.Input
    a = Mid(a, 2, 6)
ElseIf machcode = "3" Then
    a = MSComm2.Input
    a = Trim(a)
    b = 0
    For i = 1 To Len(a) - 6
        If IsNumeric(Mid(a, i, 6)) Then
            b = CLng(Mid(a, i, 6))
            comwt.Caption = b
            Exit For
        End If
    Next
ElseIf machcode = "4" Then
    a = MSComm1.Input
    a = Trim(a)
    b = 0
    For i = 1 To Len(a) - 6
        If IsNumeric(Mid(a, i, 6)) Then
            b = CLng(Mid(a, i, 6))
            comwt.Caption = b
            Exit For
        End If
    Next
ElseIf machcode = "6" Then
    a = MSComm1.Input
    a = Trim(a)
    b = 0
    For i = 1 To Len(a) - 7
        If IsNumeric(Mid(a, i, 7)) Then
            b = CLng(Mid(a, i, 7))
            comwt.Caption = b
            Exit For
        End If
    Next
ElseIf machcode = "8" Then
    a = MSComm4.Input
    
    For i = 1 To Len(a) - 7
    If LCase(Mid(a, i, 2)) = "kg" Then
        If i > 7 And Len(a) > 7 Then
            If IsNumeric(Mid(a, i - 8, 7)) Then
                b = CLng(Mid(a, i - 8, 7))
                comwt.Caption = b
            End If
        End If
        Exit For
    End If
    Next

'    For i = 1 To Len(a) - 7
'        If IsNumeric(Mid(a, i, 7)) Then
'            b = CLng(Mid(a, i, 7))
'            comwt.Caption = b
'            Exit For
'        End If
'    Next

ElseIf machcode = "7" Then
    Winsock111.GetData a
    a = Mid(a, 2, 7)
    If IsNumeric(a) Then
        comwt.Caption = Format(CDbl(a), "0000#")
    End If
Else
    If machcode <> "5" Then
        a = MSComm1.Input
        a = Mid(a, 2, 7)
    End If
End If
a = Trim(a)
If IsNumeric(a) And machcode <> "3" And machcode <> "4" And machcode <> "6" And machcode <> "7" And machcode <> "8" Then
    comwt.Caption = CLng(a)
End If

If CLng(comwt.Caption) <= 30 Then
platformin = True
End If

End Sub

Private Sub tstwt_Click()
Form1.Visible = True
End Sub


Private Sub Timer2_Timer()
On Error Resume Next
If Timer3.Enabled = True Then
    Command2.BackColor = &HFF00&
Else
    Command2.BackColor = &HFF&
End If


    MSComm3.Output = Chr$(13)
    a = MSComm3.Input
    a = Trim(a)

    If IsNumeric(Mid(a, 3, 6)) Then
        comwt.Caption = Trim(Mid(a, 3, 6))
    End If

If CLng(comwt.Caption) <= 30 Then
platformin = True
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If validtag = False Then
a = Net_Connect(unames(4), CLng(passws(4)), ips(4), CLng(paths(4)))
doorin = 1
getTag
End If

If validtag = False Then
a = Net_Connect(unames(6), CLng(passws(6)), ips(6), CLng(paths(6)))
doorin = 2
getTagg
End If
End Sub


Private Sub Timer4_Timer()
If mctr = 0 Then
    Label4.Caption = "PLEASE ENSURE THAT GATES ARE CLOSED BEFORE EVERY WEIGHMENT"
    mctr = 1
Else
    Label4.Caption = "PLATFORM SHOULD BE EMPTY BEFORE EVERY WEIGHMENT"
    mctr = 0
End If
End Sub

Private Sub traveh_Click()
unloadfrm
tag_track.Show
End Sub

Private Sub usmast_Click()
unloadfrm
Users.Show
End Sub


