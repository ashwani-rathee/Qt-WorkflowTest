VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form D_order 
   Caption         =   "DO Master"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   6840
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "D_Order.frx":0000
      Left            =   3360
      List            =   "D_Order.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   91488257
      CurrentDate     =   40649
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   91488257
      CurrentDate     =   40649
   End
   Begin VB.Frame addfrm 
      Height          =   735
      Left            =   3240
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmddel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   14880
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5010
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Selection List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   1200
      TabIndex        =   11
      Top             =   1320
      Width           =   9015
      Begin VB.Frame editfrm 
         Height          =   735
         Left            =   240
         TabIndex        =   35
         Top             =   5040
         Width           =   5175
         Begin VB.CommandButton cmdref 
            Caption         =   "&Refresh"
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
            Left            =   2760
            TabIndex        =   39
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&Cancel"
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
            Left            =   5160
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "&Save"
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
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmddel1 
            Caption         =   "&Delete"
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
            Left            =   5160
            TabIndex        =   36
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.TextBox Text4 
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
         Left            =   2160
         TabIndex        =   7
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text7 
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
         Left            =   2160
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text5 
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text6 
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
         Left            =   2160
         TabIndex        =   8
         Top             =   4080
         Width           =   1575
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
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text3 
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
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text2 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H000000FF&
         Caption         =   " New "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Kg"
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
         Left            =   3840
         TabIndex        =   34
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Ton"
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
         Left            =   3840
         TabIndex        =   33
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Product Code"
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
         Left            =   240
         TabIndex        =   31
         Top             =   3240
         Width           =   1980
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Record Type"
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
         Left            =   240
         TabIndex        =   29
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Location"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " End Date"
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
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "11 digit value required"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4320
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Order Quantity"
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
         Left            =   240
         TabIndex        =   25
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Start  Date"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Order No"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer Id"
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
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Name"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
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
      Left            =   9720
      TabIndex        =   10
      ToolTipText     =   "Unload Form"
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Order Master"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   840
      Width           =   2460
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004040&
      Height          =   615
      Left            =   1200
      TabIndex        =   18
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "D_order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opedit As Boolean
Dim recfound As Boolean
Dim str1 As String
Dim num As Integer
Dim g As Integer
Dim chkpass As Integer
Dim loaded As Boolean

Private Sub cmdadd_Click()
opedit = False
unlocktxt
addfrm.Visible = False
editfrm.Visible = True
Text2.SetFocus
End Sub

Private Sub cmdcancel_Click()
Text1.Enabled = True
opedit = False
unlocktxt
Text2.SetFocus
End Sub

Private Sub cmddel1_Click()
a = MsgBox("Want to Delete Records", vbOKCancel, "Data Deletion")
If a = 1 Then
Set rs3 = New ADODB.Recordset
rs3.Open "select * from special where tc_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount = 0 Then
Set rs5 = New ADODB.Recordset
rs5.Open "delete  from  do_master where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
MsgBox "Records Deleted Successfully", vbInformation, "Delete Confirmation"
unlocktxt
Text1.Enabled = True
Text2.SetFocus
Else
MsgBox "Record cannot be deleted, record exists in sale file"
End If

End If
End Sub

Private Sub cmdedit_Click()
opedit = True
unlocktxt
addfrm.Visible = False
editfrm.Visible = True
Text2.SetFocus

End Sub

Private Sub cmdref_Click()
unlocktxt
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub cmdsave_Click()
a = MsgBox("Want to Save Data", vbOKCancel, "Save Confirmation")

If Len(Text1.Text) = 11 Then
    If a = 1 Then
        dataset
        If recfound = False Then
            rs1.AddNew
        Else
            MsgBox "Modification not allowed"
            Exit Sub
        End If
        
        Set rs2 = New ADODB.Recordset
        rs2.Open "select * from consignee where c_code='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
        Else
            a1 = MsgBox("Customer does not exist, do you want to add to customer master", vbOKCancel, "Customer does not")
            If a1 = 1 Then
                rs2.AddNew
                rs2.Fields("c_code").Value = Trim(Text2.Text) & ""
                rs2.Fields("c_name").Value = Trim(Text3.Text) & ""
                rs2.Update
            Else
                Text2 = ""
                Text3 = ""
                Text2.SetFocus
            End If
        End If
        
        Set rs2 = New ADODB.Recordset
        rs2.Open "select * from state where state_code='" & Trim(Text5.Text) & "'", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount = 0 Then
            MsgBox "Invalid state code, please enter again"
            Text5.Text = ""
            Text5.SetFocus
        End If

        Set rs2 = New ADODB.Recordset
        rs2.Open "select * from mater where m_code='" & Trim(Text7.Text) & "'", co, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount = 0 Then
            MsgBox "Invalid state code, please enter again"
            Text7.Text = ""
            Text7.SetFocus
        End If
    
        If Trim(Text1) <> "" And Trim(Text2) <> "" And Trim(Text5) <> "" And Trim(Text6) <> "" And Trim(Combo1) <> "" And Trim(Text7) <> "" Then
            rs1.Fields("DO_NO").Value = Trim(Text1.Text)
            rs1.Fields("C_CODE").Value = Trim(Text2.Text)
            rs1.Fields("LOCATION").Value = Trim(Text5.Text) & ""
            rs1.Fields("S_DATE").Value = CDate(DTPicker1.Value)
            rs1.Fields("END_DATE").Value = CDate(DTPicker2.Value)
            rs1.Fields("o_quantity").Value = Val(Text6.Text)
            rs1.Fields("RECORD_TYPE").Value = Combo1.Text
            rs1.Fields("m_code").Value = Trim(Text7.Text)
            rs1.Update
            MsgBox "Data Updated Successfully", vbInformation, "Data Updated"
        Else
            MsgBox "All fields must be filled", , "Fields are empty"
        End If
    End If
    unlocktxt
Else
    MsgBox "Please fill all the fields properly", , "Incorrect data"
End If
Text1.Enabled = True
Text2.SetFocus

End Sub

Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
rs1.Open "select * from DO_Master where DO_NO='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
recfound = True
End If

End Sub


Private Sub cmdsave_GotFocus()
If Trim(Text1.Text) = "" Then
MsgBox "Please fill data Properly", vbInformation, "Data Error"
autoincr
Text2.SetFocus
Exit Sub
End If

If Trim(Text2.Text) = "" Then
MsgBox "Please fill data Properly", vbInformation, "Data Error"
Text2.SetFocus
Exit Sub
End If


If Trim(Text3.Text) = "" Then
MsgBox "Please fill data Properly", vbInformation, "Data Error"
Text3.SetFocus
Exit Sub
End If

If Trim(Text7.Text) = "" Then
MsgBox "Please fill data Properly", vbInformation, "Data Error"
Text7.SetFocus
Exit Sub
End If

End Sub

Private Sub Command1_Click()
achead.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub EdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub



Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1.Value + 44
End Sub



Private Sub Form_Load()
Me.Picture = main.Picture

Call Conn
unlocktxt
loaded = False
DTPicker1.Value = Date
DTPicker2.Value = DTPicker1.Value + 44
End Sub

Private Sub List2_DblClick()
If chkpass = 1 Then
    Text1.Text = Trim(padl(List2.List(List2.ListIndex), 11))
    DO_Show
    Text2.SetFocus
End If

If chkpass = 3 Then
    Text7.Text = Trim(padl(List2.List(List2.ListIndex), 6))
    Text7.SetFocus
End If

If chkpass = 4 Then
    Text5.Text = tmpcode1(List2.ListIndex)
    'Text5.Text = Trim(padl(List1.List(List1.ListIndex), 25))
    Text5.SetFocus
End If
End Sub

Private Sub txtlock()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text5.Locked = True
Text7.Locked = True
End Sub


Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text7.Text = ""
Text6.Text = "0"
Text4.Text = "0"
'Combo1.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date

Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text5.Locked = False
Text7.Locked = False
'autoincr
End Sub

Private Sub autoincr()
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

Private Sub List1_DblClick()
If chkpass = 2 Then
    Text2.Text = Left(List1.Text, 6)
    datashow
    Text2.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
chkpass = 1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
str1 = ""
num = 0
List2.Visible = True
List1.Visible = False
DO_list

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
        DO_Show
'        If opedit = True Then
'            Text1.Text = Trim(padl(List1.List(List1.ListIndex), 5))
'        End If


'        If opedit = False And recfound = True Then
'            MsgBox "This Id already Exsist ", vbInformation, "Data Duplicated"
'            Text1.Enabled = True
'            unlocktxt
'            Text2.SetFocus
'            Exit Sub
'        End If

'        If opedit = True And recfound = False Then
'            MsgBox "This Id not Exsist in Master ", vbInformation, "No Data "
'
'            Text2.SetFocus
'            Exit Sub
'        End If
        Text2.SetFocus
    End If

End If

End Sub

Private Sub Text2_Change()
'str1 = Text2.Text
''Call dd1
'
'For i = 0 To List1.ListCount - 1
'      If Trim(Text2.Text) = Left(List1.List(i), Len(Trim(Text2.Text))) Then
'                    List1.ListIndex = i
'                    Exit Sub
'        End If
'    Next i
End Sub

Private Sub Text2_GotFocus()
'chkpass = 2
'Text2.SelStart = 0
'Text2.SelLength = Len(Text2.Text)
'str1 = ""
'num = 0
'List1.Visible = True
'List2.Visible = False
'If loaded = False Then
'    Label9.Caption = "Wait, Loading Consignee List..."
'    MsgBox "Loading Consignee List can take some time..CLICK OK TO START"
'    Id_list
'End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
    If Trim(Text2.Text) <> "" Then
         datashow
         Text3.SetFocus
'        If opedit = True Then
'        Text2.Text = Trim(padl(List1.List(List1.ListIndex), 6))
'        End If
'
'        If opedit = False And recfound = True Then
'        MsgBox "This ID already exists", vbInformation, "Data Duplication"
'        Text2.Enabled = True
'        unlocktxt
'        Text2.SetFocus
'        Exit Sub
'        End If
'
'        If opedit = True And recfound = False Then
'        MsgBox "This ID does not exist in Master ", vbInformation, "No Data "
'        Text3.SetFocus
'        Exit Sub
'        End If
    End If
End If
End Sub




Private Sub product_list()
If List2.ListCount > 0 Then
    List2.Clear
    List2.Refresh
End If

Set rs = New ADODB.Recordset
    rs.Open "Select * from mater order by m_CODE", co, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    g = 0
ReDim tmp(rs.RecordCount)
While Not rs.EOF
    List2.AddItem padl(rs.Fields("m_code").Value, 6) & " " & padl(rs.Fields("m_name").Value, 50)
    rs.MoveNext
Wend
End Sub


Private Sub Id_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select c_name, c_code from consignee order by c_code", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

g = 0
'ReDim tmp(rs.RecordCount)

While Not rs.EOF
List1.AddItem padl(rs.Fields("c_code").Value, 6) & " " & padl(rs.Fields("c_name").Value, 100)
rs.MoveNext
Wend
loaded = True
Label9.Caption = "Selection List"
End Sub

Private Sub location_list()
If List2.ListCount > 0 Then
List2.Clear
List2.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from state order by state_code", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

g = 0
ReDim tmpcode1(rs.RecordCount)
While Not rs.EOF
List2.AddItem Trim(rs.Fields("state_name").Value & "")
tmpcode1(g) = rs.Fields("state_code").Value
g = g + 1
rs.MoveNext
Wend
End Sub

Private Sub DO_list()
If List2.ListCount > 0 Then
List2.Clear
List2.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from DO_Master order by DO_NO", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

g = 0
ReDim tmp(rs.RecordCount)
While Not rs.EOF
List2.AddItem padl(rs.Fields("DO_NO").Value, 11)

rs.MoveNext
Wend
End Sub


Private Sub datashow()
On Error Resume Next
recfound = False
opedit = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from consignee where c_code='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    Text1.Enabled = False
    rs1.MoveFirst
    opedit = True
    recfound = True
    Text2.Text = rs1.Fields("c_code").Value
    Text3.Text = rs1.Fields("c_name").Value
    Label16.Visible = False
Else
    Label16.Visible = True
    Text3.Text = ""
End If
End Sub

Private Sub DO_Show()
'On Error Resume Next
recfound = False
opedit = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from DO_Master where DO_NO='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    Text1.Enabled = False
    rs1.MoveFirst
    opedit = True
    recfound = True
    Text2.Text = rs1.Fields("C_CODE").Value
    Set rs2 = New ADODB.Recordset
    rs2.Open "Select * from consignee where c_code='" & Trim(Text2.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs2.RecordCount > 0 Then
        rs2.MoveFirst
        Text3.Text = rs2.Fields("c_name").Value
    End If
    Text5.Text = rs1.Fields("LOCATION").Value
    DTPicker1.Value = rs1.Fields("S_DATE").Value
    DTPicker2.Value = rs1.Fields("END_DATE").Value
    Text6.Text = rs1.Fields("O_QUANTITY").Value
    Text4.Text = CDbl(Text6.Text) / 1000
    Text7.Text = rs1.Fields("m_code").Value
    Combo1.Text = rs1.Fields("RECORD_TYPE").Value
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
Text5.SetFocus
End If

End Sub



Private Sub Text4_Change()
If IsNumeric(Text4.Text) Then
    Text6.Text = CDbl(Text4.Text) * 1000
Else
    MsgBox "Enter numeric value"
    Text4.Text = ""
    Text4.SetFocus
End If
End Sub

Private Sub Text5_Change()
str1 = Text5.Text
Call dd1
For i = 0 To List2.ListCount - 1
      If Trim(Text5.Text) = Left(List2.List(i), Len(Trim(Text2.Text))) Then
                    List2.ListIndex = i
                    Exit Sub
        End If
    Next i

End Sub

Private Sub Text5_GotFocus()
chkpass = 4
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
str1 = ""
num = 0
List1.Visible = False
List2.Visible = True
location_list

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
'KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
'Text6.SetFocus
cmdsave.SetFocus
Else
Call KeyPress1(KeyAscii)
End If
End Sub

Private Sub Text6_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
str1 = ""
num = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
'KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
DTPicker1.SetFocus
End If
End Sub

Private Sub Text7_Change()
str1 = Text7.Text
Call dd1
For i = 0 To List2.ListCount - 1
      If Trim(Text7.Text) = Left(List2.List(i), Len(Trim(Text7.Text))) Then
                    List2.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text7_GotFocus()
chkpass = 3
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
str1 = ""
num = 0
List2.Visible = True
List1.Visible = False
product_list
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
    If Trim(Text7.Text) <> "" Then
        If opedit = True Then
        Text7.Text = Trim(padl(List2.List(List2.ListIndex), 6))
        End If
        Text4.SetFocus
    End If
End If
End Sub

