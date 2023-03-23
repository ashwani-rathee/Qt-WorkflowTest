VERSION 5.00
Begin VB.Form tag_track 
   Caption         =   "TAG / VEHICLE TRACKER"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14295
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14295
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   600
      TabIndex        =   7
      Top             =   2220
      Width           =   13335
      Begin VB.TextBox Text27 
         Height          =   375
         Left            =   11280
         TabIndex        =   18
         Top             =   2700
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   11280
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   11280
         TabIndex        =   16
         Top             =   1860
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text24 
         Height          =   375
         Left            =   11280
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   11280
         TabIndex        =   14
         Top             =   1020
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   11280
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   11280
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "LAST WEIGHMENT DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2460
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TAG DETAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label5 
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
         Height          =   1755
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   12855
      End
      Begin VB.Label Label4 
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
         Height          =   1635
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   12855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8400
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
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
      Left            =   5520
      TabIndex        =   4
      Top             =   1740
      Width           =   2775
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
      Left            =   13560
      TabIndex        =   0
      ToolTipText     =   "Unload Form"
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Vehicle Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   11895
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAG / VEHICLE TRACKER"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   420
      Width           =   13575
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   6315
      Left            =   480
      Top             =   1020
      Width           =   13575
   End
End
Attribute VB_Name = "tag_track"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  tags where v_no = '" + Trim(Text1.Text) + "' order by sno desc", co1, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text21 = rs1.Fields("tagno")
Text23 = "Tag Cancellation"
Text25 = loginname
Text26 = Left(rs1.Fields("unit"), 2)
Text27 = Right(rs1.Fields("unit"), 2)

Label4.Caption = "Tag No: " + Left(rs1.Fields("tagno"), 4) + "XXXXXXXXXXXXXXXXXX" + Right(rs1.Fields("tagno"), 4) + "  Issue: " + CStr(rs1.Fields("issue")) + "  Expiry: " + CStr(rs1.Fields("expiry")) _
+ vbCrLf + "Customer: " + rs1.Fields("tc_code") + "  Vehicle: " + rs1.Fields("v_no") + "  Material: " + rs1.Fields("tm_code") _
+ vbCrLf + "DO: " + rs1.Fields("do_no") + "  Coll: " + rs1.Fields("coll_code") + "  Valid: " + rs1.Fields("valid") + "  Trips: " + CStr(rs1.Fields("tagtrips")) + "  TripsDone: " + CStr(rs1.Fields("trips_done")) _
+ vbCrLf + "Mode: " + rs1.Fields("mode") + "  Weighment type: " + rs1.Fields("wmode")
Else
    MsgBox "No active TAG found for Vehicle...it is not issued or transaction is complete"
    Exit Sub
End If
rs1.Close

Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  special where v_no = '" + Trim(Text1.Text) + "' order by sl_no desc", co1, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text22 = rs1.Fields("sl_no")
Text23 = "Transaction Cancellation"
Label5.Caption = "Serial No: " + rs1.Fields("sl_no") + "  Date in: " + CStr(rs1.Fields("date_in")) + "  Time In: " + CStr(rs1.Fields("time_in")) _
+ vbCrLf + "Operator: " + rs1.Fields("o_name") + "  First Wt: " + CStr(rs1.Fields("first_wt")) + "  Second Wt: " + CStr(rs1.Fields("second_wt")) + "  Net Wt: " + CStr(Abs(CStr(rs1.Fields("second_wt")) - CStr(rs1.Fields("first_wt")))) _
+ "  Destination: " + rs1.Fields("dest")
End If
rs1.Close

End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
On Error GoTo errm
If co1.State = 0 Then
Conn1
End If
Exit Sub
errm:
MsgBox Err.Description
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text1.Locked = False
Else
Text1.Locked = True
End If
End Sub

