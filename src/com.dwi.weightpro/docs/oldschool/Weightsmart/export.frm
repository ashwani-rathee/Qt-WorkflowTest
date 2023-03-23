VERSION 5.00
Begin VB.Form Export 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Export Data"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12030
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "X"
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
      Left            =   8880
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   6855
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   3600
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Drive"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Export Data"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000A0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000A0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   2520
      Top             =   1560
      Width           =   6855
   End
End
Attribute VB_Name = "Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim fso As New FileSystemObject
On Error GoTo err11:
a = MsgBox("Want to export daily reports", vbOKCancel, " Backup Data Base ")
If a = 1 Then

If co.State = 1 Then
co.Close
End If

str1 = App.Path + "\trans\*.*"
str2 = "z:\"
MsgBox str1 + " -> " + str2
fso.MoveFile str1, str2
'fso.CopyFile str1, str2, True
MsgBox "Data backup successful", vbInformation, "Backup Completed"
End If

err11:
If Err.Description <> "" Then
MsgBox Err.Description, vbInformation, "Data backup not completed"
End If
End Sub



Private Sub Form_Load()
Me.Picture = main.Picture

End Sub
