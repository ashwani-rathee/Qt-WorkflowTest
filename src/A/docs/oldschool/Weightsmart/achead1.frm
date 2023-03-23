VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form achead1 
   Caption         =   "A/C Head Master"
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
   Begin VB.Frame addfrm 
      Height          =   735
      Left            =   3240
      TabIndex        =   21
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   8760
      TabIndex        =   16
      Top             =   1800
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
         TabIndex        =   17
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label9 
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
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   3120
      TabIndex        =   11
      Top             =   1800
      Width           =   5655
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90439681
         CurrentDate     =   39960
      End
      Begin VB.TextBox Text4 
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
         Left            =   2160
         TabIndex        =   4
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text5 
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
         Left            =   2160
         TabIndex        =   2
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox Text3 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text2 
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
         Left            =   2160
         TabIndex        =   0
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
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
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Order Quantity"
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
         Left            =   240
         TabIndex        =   27
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dated"
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
         Left            =   240
         TabIndex        =   26
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Visible         =   0   'False
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Address"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1965
      End
   End
   Begin VB.Frame editfrm 
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   6960
      Width           =   5655
      Begin VB.CommandButton cmddel1 
         Caption         =   "&Delete"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "&Refresh"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   855
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
      Left            =   11760
      TabIndex        =   5
      ToolTipText     =   "Unload Form"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Customer Master"
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
      Left            =   6240
      TabIndex        =   19
      Top             =   1320
      Width           =   2010
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   3120
      TabIndex        =   20
      Top             =   1200
      Width           =   9135
   End
End
Attribute VB_Name = "achead1"
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
rs3.Open "select * from simple where tc_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs3.RecordCount = 0 Then
Set rs5 = New ADODB.Recordset
rs5.Open "delete  from  cust1 where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
MsgBox "Records Deleted Successfully", vbInformation, "Delete Confirmation"
unlocktxt
Text1.Enabled = True
Text2.SetFocus
Else
MsgBox "Customer cannot be deleted, record exists in sale file"
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
Text2.SetFocus

End Sub

Private Sub cmdsave_Click()
a = MsgBox("Want to Save Data", vbOKCancel, "Save Confirmation")
If a = 1 Then
dataset
If recfound = False Then
rs1.AddNew
End If

rs1.Fields("c_code").Value = Trim(Text1.Text)
rs1.Fields("c_name").Value = Trim(Text2.Text)
rs1.Fields("c_address").Value = Trim(Text3.Text) & ""
rs1.Fields("c_orderno").Value = Trim(Text5.Text) & ""
rs1.Fields("o_quantity").Value = Val(Text4.Text)
rs1.Fields("o_date").Value = CDate(DTPicker1.Value)
rs1.Update

MsgBox "Data Updated Successfully", vbInformation, "Data Updated"
End If

unlocktxt
Text1.Enabled = True
Text2.SetFocus

End Sub

Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
rs1.Open "select * from cust1 where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
recfound = True
End If

End Sub


Private Sub cmdsave_GotFocus()
If Trim(Text1.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
autoincr
Text2.SetFocus
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


'If Trim(Text4.Text) = "" Then
'MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text4.SetFocus
'Exit Sub
'End If


'If Trim(Text5.Text) = "" Then
'MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text5.SetFocus
'Exit Sub
'End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub EdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If

End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
Call Conn
unlocktxt
End Sub


Private Sub txtlock()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
End Sub


Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text4.Text = ""
DTPicker1.Value = Date

Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
autoincr
End Sub


Private Sub autoincr()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from cust1  order by val(c_code)", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveLast
Text1.Text = IncrBillNo1(rs1.Fields("c_code").Value)
Else
Text1.Text = "1"
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

Private Sub List1_DblClick()
'If opedit = True Then
    If chkpass = 1 Then
    Text1.Text = Trim(padl(List1.List(List1.ListIndex), 5))
    datashow
    Text2.SetFocus
    End If
'End If

If chkpass = 2 Then
Text1.Text = tmpcode1(List1.ListIndex)
datashow
Text2.SetFocus
End If



End Sub

Private Sub Text1_Change()
str1 = Text1.Text
Call dd1
For i = 0 To List1.ListCount - 1
      If Trim(Text1.Text) = Left(List1.List(i), Len(Trim(Text1.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text1_GotFocus()
chkpass = 1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
str1 = ""
num = 0
Id_list
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
         datashow
        If opedit = True Then
        Text1.Text = Trim(padl(List1.List(List1.ListIndex), 5))
        End If
        
        
        If opedit = False And recfound = True Then
        MsgBox "This Id already Exsist ", vbInformation, "Data Duplicated"
        Text1.Enabled = True
        unlocktxt
        Text2.SetFocus
        Exit Sub
        End If
        
        If opedit = True And recfound = False Then
        MsgBox "This Id not Exsist in Master ", vbInformation, "No Data "
        
        Text2.SetFocus
        Exit Sub
        End If
        
    Text2.SetFocus
    End If

End If

End Sub

Private Sub name_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from cust1 order by c_name", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
g = 0

End If
g = 0
ReDim tmpcode1(rs.RecordCount)
While Not rs.EOF
List1.AddItem Trim(rs.Fields("c_name").Value & "")
tmpcode1(g) = rs.Fields("c_code").Value
g = g + 1
rs.MoveNext
Wend
End Sub


Private Sub Id_list()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from CUST1 order by C_CODE", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
End If

g = 0
ReDim tmp(rs.RecordCount)
While Not rs.EOF
List1.AddItem padl(rs.Fields("c_code").Value, 5) & " " & padl(rs.Fields("c_name").Value, 25)

rs.MoveNext
Wend
End Sub

Private Sub datashow()
On Error Resume Next
recfound = False
opedit = False
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from cust1 where c_code='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
Text1.Enabled = False
rs1.MoveFirst
opedit = True
recfound = True
Text2.Text = rs1.Fields("c_name").Value
Text3.Text = rs1.Fields("c_address").Value & ""
Text5.Text = rs1.Fields("c_orderno").Value & ""
Text4.Text = rs1.Fields("o_quantity").Value
DTPicker1.Value = rs1.Fields("o_date").Value & ""
End If
End Sub

Private Sub Text2_Change()
str1 = Text2.Text
Call dd1
For i = 0 To List1.ListCount - 1
      If Trim(Text2.Text) = Left(List1.List(i), Len(Trim(Text2.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i


End Sub

Private Sub Text2_GotFocus()
chkpass = 2
name_list
If Trim(Text1.Text) = "" Then
MsgBox "Id Field Can't Be Blank ", vbInformation, "Error Data"
Exit Sub
End If
str1 = ""
num = 0
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
For i = 0 To List1.ListCount - 1
      If Trim(Text2.Text) = Left(List1.List(i), Len(Trim(Text2.Text))) Then
                    List1.ListIndex = i
                    Exit Sub
        End If
    Next i
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Call KeyPress1(KeyAscii)
End If

If KeyAscii = 13 Then
Text3.SetFocus
End If

End Sub

Private Sub Text3_Change()
str1 = Text3.Text
Call dd1
End Sub

Private Sub Text3_GotFocus()
If Trim(Text2.Text) = "" Then
Text2.SetFocus
MsgBox "Name Field Can't Be Blank ", vbInformation, "Error Data"
Exit Sub
End If
str1 = ""
num = 0
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DTPicker1.SetFocus
Else
Call KeyPress1(KeyAscii)
End If
End Sub

Private Sub Text4_Change()
str1 = Text4.Text
Call dd1
End Sub

Private Sub Text4_GotFocus()
str1 = ""
num = 0
Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
'Text6.SetFocus
cmdsave.SetFocus
Else
Call KeyPress1(KeyAscii)
End If
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
str1 = ""
num = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
'KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
DTPicker1.SetFocus
End If
End Sub

