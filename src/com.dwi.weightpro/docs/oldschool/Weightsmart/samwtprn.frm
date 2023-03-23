VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form samwtprn 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   1440
      TabIndex        =   35
      Top             =   6720
      Width           =   9015
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
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
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
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
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1695
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
      Left            =   9960
      TabIndex        =   5
      ToolTipText     =   "Unload Form"
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   9015
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
         Left            =   3480
         TabIndex        =   1
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Top             =   4200
         Width           =   1215
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
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   1320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90439681
         CurrentDate     =   39961
      End
      Begin VB.Label Label19 
         Caption         =   "Out Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label outtime 
         Caption         =   "Outtime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label odate 
         Caption         =   "Odate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   36
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label intime 
         Caption         =   "intime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
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
            Weight          =   400
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
         Top             =   240
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
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label ntw 
         Caption         =   "ntw"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   30
         Top             =   4680
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Net Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Second Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   27
         Top             =   4265
         Width           =   1335
      End
      Begin VB.Label op2 
         Caption         =   "op2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   26
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label op1 
         Caption         =   "op1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   24
         Top             =   3600
         Width           =   1590
      End
      Begin VB.Label Label9 
         Caption         =   "First Operator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label fwt 
         Caption         =   "fwt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   22
         Top             =   3840
         Width           =   1185
      End
      Begin VB.Label material 
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
         Height          =   225
         Left            =   2280
         TabIndex        =   21
         Top             =   3411
         Width           =   3405
      End
      Begin VB.Label chno 
         Caption         =   "chno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   20
         Top             =   2982
         Width           =   2085
      End
      Begin VB.Label cadr 
         Caption         =   "cadr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   19
         Top             =   2553
         Width           =   3120
      End
      Begin VB.Label cname 
         Caption         =   "cname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   18
         Top             =   2124
         Width           =   1410
      End
      Begin VB.Label vno 
         Caption         =   "vno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   17
         Top             =   1695
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "First Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   16
         Top             =   3850
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "RLW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   2990
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   2560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Vehicle No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1700
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Serial No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1270
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "In Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duplicate Weight Slip Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   480
      Width           =   3030
   End
   Begin VB.Label Label13 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   360
      Width           =   9015
   End
End
Attribute VB_Name = "samwtprn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dataset()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
End If

End Sub
Private Sub Command1_Click()
b = MsgBox("Want to Print Slip ", vbOKCancel, "Print ?")
If b = 1 Then
    printrep
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
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADC("Duplicate Weight Slip", 25) & Chr(27) & "F"
 wstream.WriteLine String(53, "-")
Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from smdtwtrep where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst

'wstream.WriteLine Chr(18) & "E" & Chr(14) & padl("Year   ", 10) & Chr(27) & "F" & Chr(18) & "E" & padl(rs1.Fields("SEASON").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("d_serial").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("1st Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 35)

If Len(rs1.Fields("O2_name").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("2nd Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O2_name").Value, 35)
End If
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Address", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("c_address").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("RLW", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("rlw").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colly/Mat Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("M_name").Value, 35)

wstream.WriteBlankLines 1
If Len(rs1.Fields("date_out").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Date Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
Else
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E"
End If

If Len(rs1.Fields("time_out").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Time Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)
Else
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E"
End If

wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 14) & padl(":", 2) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15) & Chr(27) & "F"
If rs1.Fields("second_Wt").Value > 0 Then
    wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Second Weight", 14) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15) & Chr(27) & "F"
    wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Net Weight   ", 14) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15) & Chr(27) & "F"
End If
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

Private Sub Command3_GotFocus()
If Trim(Text1.Text) = "" Then
Text1.SetFocus
Exit Sub
End If

DATAREAD
Text2.SetFocus
End Sub


Private Sub DATAREAD()
''''If checkempty1 = True Then
''''If platformin = True Then
''''Text2.Text = main.comwt.Caption
''''platformin = False
''''Else
''''MsgBox "Platform is Empty or not ready", vbCritical, "Error"
''''End If
''''Else
''''Text2.Text = main.comwt.Caption
''''End If
''''
If checkempty1 = True Then
If platformin = True Then
If Val(main.comwt.Caption) > 0 Then
If Val(Text2.Text) = 0 Then
Text2.Text = main.comwt.Caption
End If

platformin = False
Else
MsgBox "There is no Weight", vbCritical, "Weight Error"
End If

Else
MsgBox "Platform is Empty or not ready", vbCritical, " Error"
End If

Else

Text2.Text = main.comwt.Caption
End If

End Sub



Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
Conn
'username = "Rahul"
unlocktxt
DTPicker1.Value = Date
End Sub

Private Sub unlocktxt()
username = loginname
op2.Caption = username
sesson.Caption = ""
vno.Caption = ""
Text2.Text = Val(0)
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
End Sub


Private Sub datashow()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from simple where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and d_serial='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
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

If Len(rs1.Fields("o2_name").Value) > 0 Then
op2.Caption = rs1.Fields("o2_name").Value
Else
op2.Caption = "-na-"
End If

intime.Caption = rs1.Fields("time_in").Value
chno.Caption = rs1.Fields("RLW").Value

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from Mater1 where m_code='" & Trim(rs1.Fields("tm_code").Value) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    material.Caption = rs2.Fields("m_name").Value
End If
fwt.Caption = rs1.Fields("first_wt").Value

If Len(rs1.Fields("second_wt").Value) > 0 Then
    Text2.Text = rs1.Fields("second_wt").Value
    ntw.Caption = Abs(rs1.Fields("first_wt").Value - rs1.Fields("second_wt").Value)
Else
    Text2.Text = "-na-"
    ntw.Caption = "-na-"
End If
Else
    unlocktxt
    MsgBox "Serial Not Found", vbInformation, "Wrong Serial No"
End If

End Sub


Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Text1.Text) <> "" Then
datashow
'Command3.SetFocus
End If
End If

End Sub

Private Sub Text2_Change()
ntw.Caption = Abs(Val(fwt.Caption) - Val(Text2.Text))
End Sub

Private Sub Text2_GotFocus()
''''Text2.SelStart = 0
''''Text2.SelLength = Len(Text2.Text)
''''If Val(Text2.Text) = 0 Then
''''Command3.SetFocus
''''Exit Sub
''''End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text2.Text) = 0 Then
Command3.SetFocus
Else
Command1.SetFocus
End If
End If

End Sub
