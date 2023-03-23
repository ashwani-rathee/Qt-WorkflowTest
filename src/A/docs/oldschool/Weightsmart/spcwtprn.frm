VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form spcwtprn 
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14070
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   315
      Left            =   3600
      TabIndex        =   76
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   73990145
      CurrentDate     =   43061
   End
   Begin MSComCtl2.DTPicker tmpdt 
      Height          =   375
      Left            =   240
      TabIndex        =   73
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Format          =   73990145
      CurrentDate     =   41165
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   11175
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   5040
         TabIndex        =   70
         Top             =   5400
         Width           =   5895
         Begin VB.CommandButton Command6 
            Caption         =   "Resend Data"
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
            Left            =   3120
            TabIndex        =   74
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton Command5 
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
            Left            =   5160
            TabIndex        =   72
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Print"
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
            Left            =   240
            TabIndex        =   71
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   2415
         Left            =   6120
         TabIndex        =   7
         Top             =   2640
         Width           =   4695
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1800
            TabIndex        =   8
            Top             =   720
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
            Format          =   73990145
            CurrentDate     =   39961
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1800
            TabIndex        =   9
            Top             =   1080
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
            Format          =   73990145
            CurrentDate     =   39961
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal. Quantity"
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
            Left            =   120
            TabIndex        =   21
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label bqty 
            BackStyle       =   0  'Transparent
            Caption         =   "bqty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1800
            TabIndex        =   20
            Top             =   2040
            Width           =   2085
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Lifted Quantity"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lqty 
            BackStyle       =   0  'Transparent
            Caption         =   "lqty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1800
            TabIndex        =   18
            Top             =   1800
            Width           =   2085
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "DO Quantity"
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
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label oqty 
            BackStyle       =   0  'Transparent
            Caption         =   "oqty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1800
            TabIndex        =   16
            Top             =   1560
            Width           =   2085
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "End Date"
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
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
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
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Record Type"
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
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label rtype 
            BackStyle       =   0  'Transparent
            Caption         =   "rtype"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1800
            TabIndex        =   12
            Top             =   480
            Width           =   2085
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "DO Number"
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
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label dono 
            BackStyle       =   0  'Transparent
            Caption         =   "dono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            Width           =   2085
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5880
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
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8160
         TabIndex        =   22
         Top             =   240
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
         Format          =   73990145
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   2040
         TabIndex        =   75
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73990145
         CurrentDate     =   41743
      End
      Begin VB.Label outshift 
         Caption         =   "outshift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   67
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label28 
         Caption         =   "Out Shift"
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
         Left            =   6600
         TabIndex        =   66
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label inshift 
         Caption         =   "inshift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   65
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label20 
         Caption         =   "In Shift"
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
         Left            =   6600
         TabIndex        =   64
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label challandate 
         Caption         =   "challandate"
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
         Left            =   2040
         TabIndex        =   63
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label39 
         Caption         =   "Challan Date"
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
         Left            =   360
         TabIndex        =   62
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label challanno 
         Caption         =   "challanno"
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
         Left            =   2040
         TabIndex        =   61
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Challan No"
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
         Left            =   360
         TabIndex        =   60
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label collname 
         Caption         =   "collname"
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
         Left            =   2040
         TabIndex        =   59
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label29 
         Caption         =   "Coll. Name"
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
         Left            =   360
         TabIndex        =   58
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label collcode 
         Caption         =   "collcode"
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
         Left            =   2040
         TabIndex        =   57
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Coll. Code"
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
         Left            =   360
         TabIndex        =   56
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Cust Code"
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
         Left            =   360
         TabIndex        =   55
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label rlw 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   2040
         TabIndex        =   54
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Left            =   360
         TabIndex        =   53
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label cco 
         Caption         =   "cco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   52
         Top             =   1320
         Width           =   1365
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
         Height          =   255
         Left            =   6600
         TabIndex        =   51
         Top             =   1080
         Width           =   1335
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   50
         Top             =   1080
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
         Left            =   6600
         TabIndex        =   49
         Top             =   840
         Width           =   1215
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   48
         Top             =   840
         Width           =   1605
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   47
         Top             =   600
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
         Left            =   6600
         TabIndex        =   46
         Top             =   600
         Width           =   1215
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Season"
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
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ntw 
         Caption         =   "ntw"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2280
         TabIndex        =   43
         Top             =   6240
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Net Weight"
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
         Left            =   360
         TabIndex        =   42
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Second Wt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   5880
         Width           =   1815
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   40
         Top             =   2280
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   8160
         TabIndex        =   39
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "2nd Operator"
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
         Left            =   6600
         TabIndex        =   38
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "1st Operator"
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
         Left            =   6600
         TabIndex        =   37
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label fwt 
         Caption         =   "fwt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2280
         TabIndex        =   36
         Top             =   5520
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   35
         Top             =   2520
         Width           =   2925
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   34
         Top             =   2280
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   33
         Top             =   2040
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
         ForeColor       =   &H00000080&
         Height          =   465
         Left            =   2040
         TabIndex        =   32
         Top             =   1560
         Width           =   4170
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   31
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "First Wt"
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
         Left            =   360
         TabIndex        =   30
         Top             =   5520
         Width           =   1815
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
         Left            =   360
         TabIndex        =   29
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Material Code"
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
         Left            =   360
         TabIndex        =   28
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Destination"
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
         Left            =   360
         TabIndex        =   27
         Top             =   2040
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
         Left            =   360
         TabIndex        =   26
         Top             =   1560
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
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1080
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
         Left            =   360
         TabIndex        =   24
         Top             =   720
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
         Left            =   6600
         TabIndex        =   23
         Top             =   360
         Width           =   1335
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
      Left            =   12480
      TabIndex        =   3
      ToolTipText     =   "Unload Form"
      Top             =   405
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   11040
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP FTP1 
      Left            =   2280
      OleObjectBlob   =   "spcwtprn.frx":0000
      Top             =   8040
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Weighment Slip"
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
      Left            =   5880
      TabIndex        =   68
      Top             =   360
      Width           =   2235
   End
   Begin VB.Label Label13 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   1800
      TabIndex        =   69
      Top             =   240
      Width           =   11160
   End
End
Attribute VB_Name = "spcwtprn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim excdata(50) As String
Dim printop As Integer
Dim writeora As Boolean

Sub writeserver()
Dim weightord As Double
Dim weightbal As Double
Dim weightdiff As Double
Dim fctr As Integer
Dim ftime1, ftime2 As String
Dim writestr As String

ftime1 = Left(intime, 2) + Right(intime, 2) + "00"
ftime2 = Left(outtime, 2) + Right(outtime, 2) + "00"
weightord = CDbl(oqty) / 1000
weightbal = CDbl(bqty) / 1000
weightdiff = CDbl(lqty) / 1000

writeora = False
Conn3
Set rs6 = New ADODB.Recordset
rs6.Open "Select * from snm_online_data_back where unique_no='" + Trim(Text1.Text) + "'", co3, adOpenStatic, adLockOptimistic
If rs6.RecordCount > 0 Then
    MsgBox "Record already transferred"
    Exit Sub
End If
rs6.AddNew
rs6.Fields("rec_type") = rtype
rs6.Fields("do_no") = dono
rs6.Fields("validity_start_dt") = trimdate(DTPicker2.Value)
rs6.Fields("validity_end_dt") = trimdate(DTPicker3.Value)
rs6.Fields("customer_code") = CStr(cco)
rs6.Fields("coliery_code") = CStr(collcode)
rs6.Fields("product_code") = CStr(chno)
rs6.Fields("alloted_qty") = CStr(weightord)
rs6.Fields("total_lifted_qty") = CStr(weightdiff)
rs6.Fields("remaining_qty") = CStr(weightbal)
rs6.Fields("unique_no") = Text1

rs6.Fields("rlw") = CStr(rlw)
rs6.Fields("truck_no") = vno
rs6.Fields("challan_no") = challanno
rs6.Fields("challan_dt") = trimdate(DTPicker4.Value)
rs6.Fields("date_in") = trimdate(DTPicker1.Value)
rs6.Fields("time_in") = CStr(ftime1)
rs6.Fields("date_out") = trimdate(odate)
rs6.Fields("time_out") = CStr(ftime2)
rs6.Fields("gross_weight") = CStr(Text2)
rs6.Fields("tare_weight") = CStr(fwt)
rs6.Fields("net_weight") = CStr(ntw)
rs6.Fields("state_code") = CStr(cadr)
rs6.Fields("shift") = CStr(main.Label6)
rs6.Fields("login_of_operator") = CStr(op2)
DTPicker5.Value = Date
rs6.Fields("created_dt") = DTPicker5.Value
rs6.Fields("flag") = ""
rs6.Fields("mode_of_transport") = excdata(0)
rs6.Fields("other_qty") = CDbl(excdata(1))
rs6.Fields("other_amount_deposited_at_area") = CDbl(excdata(2))
rs6.Fields("total_amount_deposited") = CDbl(excdata(3))
rs6.Fields("item_code") = excdata(4)
rs6.Fields("description_of_goods") = excdata(5)
rs6.Fields("destination") = excdata(6)
rs6.Fields("type_of_coal") = excdata(7)
rs6.Fields("basic_rate") = CDbl(excdata(8))
rs6.Fields("basic_value") = CDbl(excdata(9))
rs6.Fields("add_on_price_rate") = CDbl(excdata(10))
rs6.Fields("add_on_price_value") = CDbl(excdata(11))
rs6.Fields("selective_loading_charge_rate") = CDbl(excdata(12))
rs6.Fields("selective_loading_charge_value") = CDbl(excdata(13))
rs6.Fields("weighment_charge_rate") = CDbl(excdata(14))
rs6.Fields("weighment_charge_value") = CDbl(excdata(15))
rs6.Fields("surface_transport_charge_rate") = CDbl(excdata(16))
rs6.Fields("surface_transport_charge_value") = CDbl(excdata(17))
rs6.Fields("excisable_gross_value") = CDbl(excdata(18))
rs6.Fields("central_excise_duty_rate") = CDbl(excdata(19))
rs6.Fields("central_excise_duty_value") = CDbl(excdata(20))
rs6.Fields("edu_cess_rate") = CDbl(excdata(21))
rs6.Fields("edu_cess_value") = CDbl(excdata(22))
rs6.Fields("sh_cess_on_excise_rate") = CDbl(excdata(23))
rs6.Fields("sh_cess_on_excise_value") = CDbl(excdata(24))
rs6.Fields("total_excise_duty") = CDbl(excdata(25))
rs6.Fields("royalty_rate") = CDbl(excdata(26))
rs6.Fields("royalty") = CDbl(excdata(27))
rs6.Fields("stowing_excise_duty_rate") = CDbl(excdata(28))
rs6.Fields("stowing_excise_duty_value") = CDbl(excdata(29))
rs6.Fields("clean_energy_cess_rate") = CDbl(excdata(30))
rs6.Fields("clean_energy_cess_value") = CDbl(excdata(31))
rs6.Fields("road_cess_rate") = CDbl(excdata(32))
rs6.Fields("road_cess_value") = CDbl(excdata(33))
rs6.Fields("pwd_cess_rate") = CDbl(excdata(34))
rs6.Fields("pwd_cess_value") = CDbl(excdata(35))
rs6.Fields("ambh_cess_rate") = CDbl(excdata(36))
rs6.Fields("ambh_cess_value") = CDbl(excdata(37))
rs6.Fields("other_charges_rate") = CDbl(excdata(38))
rs6.Fields("other_charges_value") = CDbl(excdata(39))
rs6.Fields("total_value") = CDbl(excdata(40))
rs6.Fields("vat_cst_rate") = CDbl(excdata(41))
rs6.Fields("vat_cst_value") = CDbl(excdata(42))
rs6.Fields("gross_value") = CDbl(excdata(43))
rs6.Fields("tax_collected_at_source_rate") = CDbl(excdata(44))
rs6.Fields("tax_collected_at_source_value") = CDbl(excdata(45))
rs6.Fields("gross_bill_value") = CDbl(excdata(46))
rs6.Fields("bill_number") = excdata(47)
rs6.Fields("bill_date") = excdata(48)
rs6.Fields("bill_type") = padl(excdata(49), 2)
rs6.Fields("approval_flag") = ""
rs6.Fields("approved_by") = ""
DTPicker5.Value = Date
rs6.Fields("approval_date") = DTPicker5.Value
rs6.Fields("bill_flag") = ""
rs6.Update
writeora = True
rs6.Close
co3.Close
End Sub



Private Sub dataset()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from special where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# and sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
End If
End Sub

'Function trimslash(tda As String) As String
'trimslash = Mid(tda, 1, 2) + Mid(tda, 4, 2) + Mid(tda, 7, 4)
'End Function

Function trimdate(tda As Date) As String
trimdate = Format(Day(tda), "0#") & Format(Month(tda), "0#") & Format(Year(tda), "0###")
End Function


Function writefile() As Boolean
Dim fname As String
Dim weightord As Double
Dim weightbal As Double
Dim weightdiff As Double
Dim fctr As Integer
Dim ftime1, ftime2 As String
Dim writestr As String

ftime1 = Left(intime, 2) + Right(intime, 2) + "00"
ftime2 = Left(outtime, 2) + Right(outtime, 2) + "00"
weightord = CDbl(oqty) / 1000
weightbal = CDbl(bqty) / 1000
weightdiff = CDbl(lqty) / 1000
fname = "gr" + Trim(dono) + Trim(Text1.Text) + ".txt"

writestr = padl(Trim(rtype), 1) & "|" & padl(Trim(dono), 11) & "|" & padl(trimdate(DTPicker2.Value), 8) & "|" & padl(trimdate(DTPicker3.Value), 8) & "|" & padl(Trim(cco), 6) & "|" & padl(Trim(collcode), 4) & "|" & padl(Trim(chno), 10) & "|" & padl(Format(weightord, "0####.##0"), 9) & "|" & padl(Format(Trim(CStr(weightdiff)), "0####.##0"), 9) & "|" & padl(Format(weightbal, "0####.##0"), 9) & "|" & padl(Trim(Text1), 12) & "|" & padl(Format(CDbl(rlw) / 1000, "0####.##0"), 9) & "|" & padl(Trim(vno), 15) & "|" & padl(Trim(challanno.Caption), 8) & "|" & padl(trimdate(DTPicker4.Value), 8) & "|" & padl(trimdate(DTPicker1.Value), 8) & "|" & padl(Trim(ftime1), 6) & "|" & padl(trimdate(odate), 8) & "|" & padl(Trim(ftime2), 6) & "|" & padl(Format(CDbl(Text2) / 1000, "0#####.##0"), 10) & "|" & padl(Format(CDbl(fwt) / 1000, "0#####.##0"), 10) & "|" & padl(Format(CDbl(ntw) / 1000, "0#####.##0"), 10) & "|" & padl(Trim(cadr), 2) & "|" & padl(Trim(main.Label6), 1) & "|" & padl(Trim(op2), 6) & "|" _
 + padl(excdata(0), 4) + "|" + padl(excdata(1), 10) + "|" + padl(excdata(2), 13) + "|" + padl(excdata(3), 13) + "|" + padl(excdata(4), 8) + "|" + padl(excdata(5), 8) + "|" + padl(excdata(6), 20) + "|" + padl(excdata(7), 2) + "|" + padl(excdata(8), 8) + "|" + padl(excdata(9), 13) + "|" + padl(excdata(10), 8) + "|" + padl(excdata(11), 13) + "|" + padl(excdata(12), 8) + "|" + padl(excdata(13), 13) + "|" + padl(excdata(14), 8) + "|" + padl(excdata(15), 13) + "|" + padl(excdata(16), 8) + "|" + padl(excdata(17), 13) + "|" + padl(excdata(18), 13) + "|" + padl(excdata(19), 8) + "|" + padl(excdata(20), 13) + "|" + padl(excdata(21), 8) + "|" + padl(excdata(22), 13) + "|" + padl(excdata(23), 8) + "|" + padl(excdata(24), 13) + "|" _
 + padl(excdata(25), 13) + "|" + padl(excdata(26), 8) + "|" + padl(excdata(27), 13) + "|" + padl(excdata(28), 8) + "|" + padl(excdata(29), 13) + "|" + padl(excdata(30), 8) + "|" + padl(excdata(31), 13) + "|" + padl(excdata(32), 8) + "|" + padl(excdata(33), 13) + "|" + padl(excdata(34), 8) + "|" + padl(excdata(35), 13) + "|" + padl(excdata(36), 8) + "|" + padl(excdata(37), 13) + "|" + padl(excdata(38), 8) + "|" + padl(excdata(39), 13) + "|" + padl(excdata(40), 13) + "|" + padl(excdata(41), 8) + "|" + padl(excdata(42), 13) + "|" + padl(excdata(43), 13) + "|" + padl(excdata(44), 8) + "|" + padl(excdata(45), 13) + "|" + padl(excdata(46), 13) + "|" + padl(excdata(47), 10) + "|" + padl(excdata(48), 8) + "|" + padl(excdata(49), 2) + "|" + vbCrLf

writefile = True

If FTP1.Connect Then
    If FTP1.ChangeRemoteDir(paths(2)) Then
        If FTP1.PutFileFromTextData(fname, writestr) Then
            writefile = True
            If FTP1.GetRemoteFileTextData(fname) = writestr Then
                MsgBox "File transfered fo Headquarter, Verification successful"
                writefile = True
            Else
                MsgBox "File verification failed..check connection and save again"
                writefile = False
            End If
        Else
            MsgBox "cound not write file..check connection and save again"
            writefile = False
        End If
    Else
        MsgBox "Remote directory not found..check connection and save again"
        writefile = False
    End If
Else
    MsgBox "No connection to headquarter..check connection and save again"
    writefile = False
End If

End Function


Private Sub printrep()
Dim ordst As String
Dim orden As String
On Error Resume Next
spcsndwt.Sections(1).Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
spcsndwt.Sections(1).Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
spcsndwt.Sections(1).Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)
spcsndwt.Sections(1).Controls("Label4").Caption = PADC("Duplicate Weighment Slip", 27)

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set spcsndwt.DataSource = rs1
spcsndwt.Sections(1).Controls("Text1").DataField = "SEASON"
Set spcsndwt.Sections(1).Controls("Image1").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_sw1.jpg")
Set spcsndwt.Sections(1).Controls("Image2").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_sw2.jpg")
Set spcsndwt.Sections(1).Controls("Image3").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_sw3.jpg")
Set spcsndwt.Sections(1).Controls("Image4").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_sw4.jpg")
spcsndwt.Sections(1).Controls("Label5").Caption = padl("Year   ", 20) & padl(":", 2) & padl(rs1.Fields("SEASON").Value, 15)
spcsndwt.Sections(1).Controls("Label6").Caption = padl("Date   ", 20) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15)
spcsndwt.Sections(1).Controls("Label7").Caption = padl("Time In", 20) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15)
spcsndwt.Sections(1).Controls("Label8").Caption = PADR("Daily Serial", 20) & padl(":", 2) & padl(rs1.Fields("sl_no").Value, 15)

spcsndwt.Sections(1).Controls("Label9").Caption = "TAG No : " & Left(rs1.Fields("tag").Value, 4) & "XXXXXXXXXX" & Right(rs1.Fields("tag").Value, 4)
spcsndwt.Sections(1).Controls("Label10").Caption = PADR("Operator 1 ", 12) & padl(": ", 2) & rs1.Fields("O_name").Value & PADR("      Operator 2 ", 12) & padl(": ", 2) & rs1.Fields("O2_name").Value

spcsndwt.Sections(1).Controls("Label11").Caption = PADR("Vehicle No", 20) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 15)

spcsndwt.Sections(1).Controls("Label12").Caption = PADR("Name", 20) & Chr(27) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
spcsndwt.Sections(1).Controls("Label13").Caption = PADR("Destination", 20) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)
spcsndwt.Sections(1).Controls("Label14").Caption = PADR("Order No.", 20) & padl(":", 2) & padl(rs1.Fields("DO_NO").Value, 35)

ordst = Mid(rs1.Fields("do_start_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 1, 4)
orden = Mid(rs1.Fields("do_end_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 1, 4)
spcsndwt.Sections(1).Controls("Label15").Caption = PADR("Order Validity", 20) & padl(":", 2) & padl(ordst, 10) & " - " & padl(orden, 10)

spcsndwt.Sections(1).Controls("Label16").Caption = PADR("Colliery Name", 20) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
spcsndwt.Sections(1).Controls("Label17").Caption = PADR("Material", 20) & padl(":", 2) & padl(chno.Caption, 35)
spcsndwt.Sections(1).Controls("Label18").Caption = PADR("Challan No", 20) & padl(":", 2) & padl(rs1.Fields("Challan_No").Value, 15)

spcsndwt.Sections(1).Controls("Label19").Caption = "Order Qty: " & oqty & "   Lifted Qty: " & lqty & "   Balance Qty: " & bqty
spcsndwt.Sections(1).Controls("Label20").Caption = padl("Date In", 20) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & padl("Date Out", 10) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
spcsndwt.Sections(1).Controls("Label21").Caption = padl("Time In", 20) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & padl("Time Out", 10) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)

spcsndwt.Sections(1).Controls("Label22").Caption = PADR("RLW", 20) & padl(":", 2) & padl(rs1.Fields("RLW").Value, 15) & "Kg"

spcsndwt.Sections(1).Controls("Label23").Caption = PADR("First Weight", 14) & padl(str(rs1.Fields("First_Wt").Value) + " Kg", 35) & Chr(27)

If Val(Text2.Text) > 0 Then
spcsndwt.Sections(1).Controls("Label24").Caption = PADR("Second Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15)
spcsndwt.Sections(1).Controls("Label25").Caption = PADR("Net Weight   ", 13) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15)
End If

spcsndwt.Sections(1).Controls("Label26").Caption = PADR(COMPAUTH, 20) & padl(" ", 2) & padl("Weighing Operator", 30)

'spcsndwt.Sections(1).Controls("Label20").Caption = padl(compdes, 100)
Else
spcsndwt.Sections(1).Controls("Label5").Caption = PADC(Trim("No Such Serial No Created"), 30)
End If

If printop = 1 Then
spcsndwt.Show vbModal
End If
End Sub


Private Sub printdata()
Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
Dim osdat As String
Dim oedat As String
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("Duplicate Weighment Slip", 28) & Chr(27) & "F"
 wstream.WriteLine Chr(18) & String(53, "-")
Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst


'wstream.WriteLine Chr(18) & "E" & Chr(14) & padl("Year   ", 10) & Chr(27) & "F" & Chr(18) & "E" & padl(rs1.Fields("SEASON").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("SL_NO").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("1st Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 35)
If Len(rs1.Fields("O2_name").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("2nd Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O2_name").Value, 35)
End If
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name    ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Destination ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("do_no").Value, 35)

osdat = Mid(rs1.Fields("do_start_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 1, 4)
oedat = Mid(rs1.Fields("do_end_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 1, 4)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order Validity", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(osdat, 35) & " - " & padl(oedat, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colliery Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Material", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(chno.Caption, 35)
If Len(rs1.Fields("Challan_No").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Challan No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Challan_No").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & "Order Qty: " & oqty & "     Lifted Qty: " & lqty & "     Balance Qty: " & bqty
End If
wstream.WriteBlankLines 1

If Len(rs1.Fields("O2_name").Value) > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Date Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Time Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)
Else
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E"
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E"
End If
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("RLW", 13) & padl(":", 2) & padl(str(rs1.Fields("RLW").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15) & Chr(27) & "F"
If rs1.Fields("second_Wt").Value > 0 Then
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Second Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Net Weight   ", 13) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15) & Chr(27) & "F"
End If
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & String(53, "-")
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR(COMPAUTH, 20) & Chr(27) & "F" & padl(":", 2) & padl("Weighing Operator", 20)
wstream.WriteLine Chr(18) & String(53, "-")
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



Private Sub Command2_Click()
unlocktxt
End Sub




Private Sub DATAREAD()
If checkempty1 = True Then
If platformin = True Then

    If Val(Text2.Text) = 0 Then
        Text2.Text = main.comwt.Caption
    End If

    platformin = False


Else
MsgBox "Platform is Empty or not ready", vbCritical, "Error"
End If
Else
Text2.Text = main.comwt.Caption
End If
End Sub





Private Sub Command3_Click()
    b = 1
    While b = 1
    b = MsgBox("Print Slip ", vbOKCancel, "Print ?")
    If b = 1 Then
'        If camcode = 0 Then
'            printdata
'            dosprint
'        Else
            printop = 1
            printrep
'        End If
    End If
    Wend
    
    b = 1
    
    If Len(op2.Caption) > 0 And Val(Text2.Text) > 0 Then
        While b = 1
        b = MsgBox("Print Invoice ?", vbOKCancel, "Print ?")
        If b = 1 Then
'            printinv
'            dosprint
            printop = 1
            invoicerep
        End If
        Wend
    End If
End Sub


Private Sub printinv()
Dim excgross As Double
Dim excduty As Double
Dim excdut1 As Long
Dim totgross As Double
Dim decpart As Double
Dim lifted As Double
Dim dd1amt As Double
Dim dd2amt As Double
Dim dd3amt As Double
Dim othamt As Double
Dim dsno As Long
Dim billfnd As Boolean

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cadata where do_no='" & Trim(dono.Caption) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
Else
    MsgBox "Relevant CA Data not found"
    Exit Sub
End If

Set rs4 = New ADODB.Recordset
rs4.Open "Select * from excisedata", co, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
    rs4.MoveFirst
Else
    MsgBox "Relevant Excise Data not found"
    Exit Sub
End If

excgross = 0
excduty = 0
totgross = 0

Set rs5 = New ADODB.Recordset
rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' and sl_no = '" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs5.RecordCount = 0 Then
    rs5.Close
    Set rs5 = New ADODB.Recordset
    rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' order by billno", co, adOpenKeyset, adLockOptimistic
    If rs5.RecordCount > 0 Then
        rs5.MoveLast
        dsno = rs5.Fields("billno").Value + 1
    Else
        dsno = 1
    End If

    rs5.AddNew
    rs5.Fields(0).Value = Trim(sesson.Caption)
    rs5.Fields(1).Value = dsno
    rs5.Fields(2).Value = Trim(Text1.Text)
    rs5.Update
Else
    dsno = rs5.Fields("billno").Value
End If
rs5.Close

Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine PADC("Original/ Duplicate/ Triplicate/ Quadruplicate", 80)
wstream.WriteLine String(80, "-")
wstream.WriteLine String(50, " ") & PADR("for Road", 30)
wstream.WriteLine PADR("           BHARAT COKING COAL LIMITED             ", 50) & PADR("Pre Authenticated", 30)
wstream.WriteLine PADR("       A Subsidiary of Coal India Limited         ", 50)
wstream.WriteLine PADR("       Regd. Office: Koyla Bhavan, DHANBAD        ", 50) & PADR("for Bharat Coking Coal Limited", 30)
wstream.WriteLine String(50, " ") & PADR("Authorized Signatory", 30)
wstream.WriteLine String(80, "-")
wstream.WriteLine PADC("Tax Invoice under rule 11 of C/E Rules 2002", 80)
wstream.WriteLine "Type of Customer: " & rtype.Caption
wstream.WriteLine "Bill No: " & "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6) & "       Bill Date: " & odate.Caption
wstream.WriteLine String(80, "-")

wstream.WriteLine PADR("Unit Name: ", 16) & PADR(collname, 23) & PADR(" Purchaser Name: ", 17) & PADR(cname, 24)
wstream.WriteLine PADR("Address: ", 16) & PADR(Trim(main.Label3.Caption), 23) & PADR(" Address: ", 17) & PADR(rs2.Fields("destination"), 24)
wstream.WriteLine PADR("Sel.Exc.Reg.No.: ", 16) & PADR(rs4.Fields("u_excno"), 23) & PADR(" Pur.Exc.RegNo.: ", 17) & PADR(rs2.Fields("exc_reg_no"), 24)
wstream.WriteLine PADR("Range: ", 16) & PADR(rs4.Fields("u_range"), 23) & PADR(" Range: ", 17) & PADR(rs2.Fields("range"), 24)
wstream.WriteLine PADR("Division: ", 16) & PADR(rs4.Fields("u_division"), 23) & PADR(" Division: ", 17) & PADR(rs2.Fields("division"), 24)
wstream.WriteLine PADR("Commissionerate: ", 16) & PADR(rs4.Fields("u_commissionrate"), 23) & PADR(" Commissionerate: ", 17) & PADR(rs2.Fields("commissionerate"), 24)
wstream.WriteLine PADR("Vat Tin No.: ", 16) & PADR(rs4.Fields("u_tinno"), 23) & PADR(" Vat Tin No.: ", 17) & PADR(rs2.Fields("vat_tin_no"), 24)
wstream.WriteLine PADR("CST No.: ", 16) & PADR(rs4.Fields("u_cstno"), 23) & PADR(" CST No.: ", 17) & PADR(rs2.Fields("cst_no"), 24)
wstream.WriteLine String(80, "-")
wstream.WriteLine PADC("MODE OF", 10) & "|" & PADC("APPLICATION", 20) & "|" & PADC("DO QTY", 8) & "|" & PADC("DRAFT", 12) & "|" & PADC("DRAFT", 10) & "|" & PADC("DRAFT", 15)
wstream.WriteLine PADC("TRANSPORT", 10) & "|" & PADC("NUMBER", 11) & "|" & PADC("DATE", 8) & "|" & PADC("", 8) & "|" & PADC("NUMBER", 12) & "|" & PADC("DATE", 10) & "|" & PADC("AMOUNT", 15)
wstream.WriteLine String(80, "-")

dd1amt = rs2.Fields("draft_amt1")
dd2amt = rs2.Fields("draft_amt2")
dd3amt = rs2.Fields("draft_amt3")
othamt = 0

wstream.WriteLine PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no1"), 12) & "|" & PADC(rs2.Fields("draft_dt1"), 10) & "|" & padl(Format(dd1amt, "0#######.#0"), 15)
wstream.WriteLine PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no2"), 12) & "|" & PADC(rs2.Fields("draft_dt2"), 10) & "|" & padl(Format(dd2amt, "0#######.#0"), 15)
wstream.WriteLine PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no3"), 12) & "|" & PADC(rs2.Fields("draft_dt3"), 10) & "|" & padl(Format(dd3amt, "0#######.#0"), 15)
wstream.WriteLine PADR("OTHER QTY/AMOUNT:", 18) & PADR("", 20) & " Other amount deposited : " & padl(Format(othamt, "0######0.#0"), 15)
wstream.WriteLine String(80, "-")
wstream.WriteLine PADR("", 38) & "  Total amount deposited : " & padl(Format(dd1amt + dd2amt + dd3amt + othamt, "0######0.#0"), 15)
wstream.WriteLine String(80, "-")
wstream.WriteLine "DO.No.: " & PADR(dono, 11) & " Date: " & rs2.Fields("do_date") & " Grade: " & rs2.Fields("grade") & " DO Qty : " & rs2.Fields("do_qty")
wstream.WriteLine String(80, "-")
wstream.WriteLine PADC("Loading Date", 15) & PADC("Challan No", 15) & PADC("Truck No", 15) & PADC("Grade", 15) & PADC("Qty Lifted", 15)
wstream.WriteLine String(80, "-")
lifted = CDbl(ntw.Caption) / 1000
wstream.WriteLine PADC(odate, 15) & PADC(challanno.Caption, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
wstream.WriteLine String(80, "-")
wstream.WriteLine String(25, " ") & PADR("BILLED HEAD", 30) & padl("RATE", 12) & padl("VALUE", 13)
wstream.WriteLine String(25, " ") & PADR("", 30) & String(25, "-")
wstream.WriteLine String(25, " ") & PADR("Basic Value", 30) & padl(Format(CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 12) & padl(Format(lifted * CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & PADR("Add on Price (WRC)", 30) & padl(CDbl(rs2.Fields("wrc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("wrc")), "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & PADR("Selective Loading Charge", 30) & padl(CDbl(rs2.Fields("slc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("slc")), "#####0.#0"), 13)
wstream.WriteLine PADR("Item Code: 27011910", 25) & PADR("Weighment Charge", 30) & padl(CDbl(rs2.Fields("weighment_chg")), 12) & padl(Format(lifted * CDbl(rs2.Fields("weighment_chg")), "#####0.#0"), 13)
wstream.WriteLine PADR("Type of Coal:", 25) & PADR("Royalty", 30) & PADR(CDbl(rs2.Fields("royalty")), 12) & padl(Format(lifted * CDbl(rs2.Fields("royalty")), "#####0.#0"), 13)
wstream.WriteLine PADR("Non-Coking Coal", 25) & PADR("Stowing Excise Duty(SED)", 30) & padl(CDbl(rs2.Fields("sed")), 12) & padl(Format(lifted * CDbl(rs2.Fields("sed")), "#####0.#0"), 13)

excgross = (lifted * CDbl(rs2.Fields("basic_rate"))) + (lifted * CDbl(rs2.Fields("wrc"))) + (lifted * CDbl(rs2.Fields("slc"))) + (lifted * CDbl(rs2.Fields("weighment_chg"))) + lifted * CDbl(rs2.Fields("royalty")) + lifted * CDbl(rs2.Fields("sed"))
wstream.WriteLine PADR("Description of Goods", 30) & PADR("Excisable Gross", 32) & padl(Format(excgross, "#####0.#0"), 13)
wstream.WriteLine PADR("Raw Coal", 30) & PADR("Cent Excise @" + rs2.Fields("cent_exc_rate") + "%", 32) & padl(Format(excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01, "#####0.#0"), 13)
wstream.WriteLine PADR("Destination: " & Trim(cadr), 30) & PADR("Ed Cess on Excise @" + rs2.Fields("edu_cess_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01, "#####0.#0"), 13)
wstream.WriteLine PADR(" ", 30) & PADR("S/H Cess on Excise @" + rs2.Fields("high_edu_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01, "#####0.#0"), 13)
excduty = (excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01)
totgross = excgross + excduty
wstream.WriteLine PADR("", 25) & PADR("Total Excise Duty", 42) & padl(Format(excduty, "#####0.#0"), 13)
decpart = Right(Format(excduty, "######.#0"), 2)
excdut1 = Left(Format(excduty, "0#####.#0"), Len(Format(excduty, "0#####.#0")) - 3)
wstream.WriteLine "Rupees " & worda(Val(excdut1)) & " and " & worda(CLng(decpart)) & " paise only"
wstream.WriteLine String(80, "-")
wstream.WriteLine String(25, " ") & PADR("Clean Energy Cess", 30) & padl(CDbl(rs2.Fields("clean_engy_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("clean_engy_cess")), "#####0.#0"), 13)
totgross = totgross + lifted * CDbl(rs2.Fields("clean_engy_cess"))
wstream.WriteLine String(25, " ") & PADR("Road/ RE Cess", 30) & padl(CDbl(rs2.Fields("road_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("road_cess")), "#####0.#0"), 13)
totgross = totgross + lifted * CDbl(rs2.Fields("road_cess"))
wstream.WriteLine String(25, " ") & PADR("PWD Cess/ MADA(Bazaar)", 30) & padl(CDbl(rs2.Fields("bazar_fee")), 12) & padl(Format(lifted * CDbl(rs2.Fields("bazar_fee")), "#####0.#0"), 13)
totgross = totgross + lifted * CDbl(rs2.Fields("bazar_fee"))
wstream.WriteLine String(25, " ") & PADR("AMBH Cess", 30) & padl(CDbl(rs2.Fields("ambh_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("ambh_cess")), "#####0.#0"), 13)
totgross = totgross + lifted * CDbl(rs2.Fields("ambh_cess"))
wstream.WriteLine String(25, " ") & PADR("Other Charges", 30) & padl(" ", 12) & padl(CDbl(rs2.Fields("other_charges")), 13)
totgross = totgross + lifted * CDbl(rs2.Fields("other_charges"))
wstream.WriteLine String(25, " ") & String(55, "-")

wstream.WriteLine String(25, " ") & PADR("TOTAL VALUE", 42) & padl(Format(totgross, "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & PADR("VAT/CST@ " + rs2.Fields("tax_percent") + "%", 30) & padl("JST", 12) & padl(Format(totgross * 0.01 * CDbl(rs2.Fields("tax_percent")), "#####0.#0"), 13)

If IsNull(CDbl(rs2.Fields("tcs"))) Then rs2.Fields("tcs") = 0
tcs = totgross * 0.01 * CDbl(rs2.Fields("tcs"))

totgross = totgross + (totgross * 0.01 * CDbl(rs2.Fields("tax_percent")))
wstream.WriteLine String(25, " ") & PADR("TCS@ " & rs2.Fields("tcs"), 30) & padl(" ", 12) & padl(Format(tcs, "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & String(55, "-")
totgross = totgross + tcs

wstream.WriteLine String(25, " ") & PADR("Gross Value", 42) & padl(Format(totgross, "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & String(55, "-")
wstream.WriteLine "Gross value in words:"
decpart = Right(Format(totgross, "######.#0"), 2)
totgross = Left(Format(totgross, "######.#0"), Len(Format(totgross, "######.#0")) - 3)
wstream.WriteLine "Rupees " & worda(Val(totgross)) & " and " & worda(CLng(decpart)) & " paise only"
wstream.WriteLine String(80, "-")
wstream.WriteLine String(80, " ")
wstream.WriteLine PADR("Prepared By", 26) & PADR("Checked By", 26) & PADR("Auth. Signatory", 26)
wstream.WriteLine PADR(" ", 45) & PADR("Under Jurisdiction of Jharkhand", 35)
wstream.WriteBlankLines 1
wstream.Close
Set wstream = Nothing
rs2.Close
rs4.Close
End Sub


Private Sub invoicerep()
Dim excgross As Double
Dim excduty As Double
Dim excdut1 As Long
Dim totgross As Double
Dim decpart As Double
Dim lifted As Double
Dim dd1amt As Double
Dim dd2amt As Double
Dim dd3amt As Double
Dim othamt As Double
Dim tcs As Double
Dim dsno As Long
Dim billfnd As Boolean

excdata(0) = "ROAD"
excdata(1) = "0"
excdata(2) = "0"
excdata(3) = "0"
excdata(4) = "27011910"
excdata(5) = "RAW COAL"
excdata(7) = "NC"

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cadata where do_no='" & Trim(dono.Caption) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
Else
    MsgBox "Relevant CA Data not found"
    Exit Sub
End If

Set rs4 = New ADODB.Recordset
rs4.Open "Select * from excisedata", co, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
    rs4.MoveFirst
Else
    MsgBox "Relevant Excise Data not found"
    Exit Sub
End If

excgross = 0
excduty = 0
totgross = 0

Set rs5 = New ADODB.Recordset
rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' and sl_no = '" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs5.RecordCount = 0 Then
    rs5.Close
    Set rs5 = New ADODB.Recordset
    rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' order by billno", co, adOpenKeyset, adLockOptimistic
    If rs5.RecordCount > 0 Then
        rs5.MoveLast
        dsno = rs5.Fields("billno").Value + 1
    Else
        dsno = 1
    End If

    rs5.AddNew
    rs5.Fields(0).Value = Trim(sesson.Caption)
    rs5.Fields(1).Value = dsno
    rs5.Fields(2).Value = Trim(Text1.Text)
    rs5.Update
Else
    dsno = rs5.Fields("billno").Value
End If
rs5.Close

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set invrep.DataSource = rs1
invrep.Sections("Section1").Controls("Text1").DataField = "SEASON"
End If

invrep.Sections("Section1").Controls("Label1").Caption = "Original/ Duplicate/ Triplicate/ Quadruplicate"
invrep.Sections("Section1").Controls("Label2").Caption = "for Road"

invrep.Sections("Section1").Controls("Label3").Caption = PADR("           BHARAT COKING COAL LIMITED             ", 50) & PADR("Pre Authenticated", 30)
invrep.Sections("Section1").Controls("Label4").Caption = PADR("       A Subsidiary of Coal India Limited         ", 50)
invrep.Sections("Section1").Controls("Label5").Caption = PADR("       Regd. Office: Koyla Bhavan, DHANBAD        ", 50) & PADR("for Bharat Coking Coal Limited", 30)
invrep.Sections("Section1").Controls("Label6").Caption = String(50, " ") & PADR("Authorized Signatory", 30)
invrep.Sections("Section1").Controls("Label7").Caption = PADC("Tax Invoice under rule 11 of C/E Rules 2002", 80)
invrep.Sections("Section1").Controls("Label8").Caption = "Type of Customer: " & rtype.Caption
invrep.Sections("Section1").Controls("Label9").Caption = "Bill No: " & "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6) & "       Bill Date: " & odate.Caption
invrep.Sections("Section1").Controls("Label10").Caption = "TAG No : " & Left(rs1.Fields("tag").Value, 4) & "XXXXXXXXXX" & Right(rs1.Fields("tag").Value, 4)

invrep.Sections("Section1").Controls("Label11").Caption = PADR("Unit Name: ", 16) & PADR(collname, 23) & PADR(" Purchaser Name: ", 17) & PADR(cname, 24)
invrep.Sections("Section1").Controls("Label12").Caption = PADR("Address: ", 16) & PADR(Trim(main.Label3.Caption), 23) & PADR(" Address: ", 17) & PADR(rs2.Fields("destination"), 24)
excdata(6) = rs2.Fields("destination")
invrep.Sections("Section1").Controls("Label13").Caption = PADR("Sel.Exc.Reg.No.: ", 16) & PADR(rs4.Fields("u_excno"), 23) & PADR(" Pur.Exc.RegNo.: ", 17) & PADR(rs2.Fields("exc_reg_no"), 24)
invrep.Sections("Section1").Controls("Label14").Caption = PADR("Range: ", 16) & PADR(rs4.Fields("u_range"), 23) & PADR(" Range: ", 17) & PADR(rs2.Fields("range"), 24)
invrep.Sections("Section1").Controls("Label15").Caption = PADR("Division: ", 16) & PADR(rs4.Fields("u_division"), 23) & PADR(" Division: ", 17) & PADR(rs2.Fields("division"), 24)
invrep.Sections("Section1").Controls("Label16").Caption = PADR("Commissionerate: ", 16) & PADR(rs4.Fields("u_commissionrate"), 23) & PADR(" Commissionerate: ", 17) & PADR(rs2.Fields("commissionerate"), 24)
invrep.Sections("Section1").Controls("Label17").Caption = PADR("Vat Tin No.: ", 16) & PADR(rs4.Fields("u_tinno"), 23) & PADR(" Vat Tin No.: ", 17) & PADR(rs2.Fields("vat_tin_no"), 24)
invrep.Sections("Section1").Controls("Label18").Caption = PADR("CST No.: ", 16) & PADR(rs4.Fields("u_cstno"), 23) & PADR(" CST No.: ", 17) & PADR(rs2.Fields("cst_no"), 24)
invrep.Sections("Section1").Controls("Label19").Caption = PADR("PAN No.: ", 16) & PADR(rs4.Fields("u_panno"), 23) & PADR(" PAN No.: ", 17) & PADR(rs2.Fields("PAN"), 24)
invrep.Sections("Section1").Controls("Label20").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label21").Caption = PADC("MODE OF", 10) & "|" & PADC("APPLICATION", 20) & "|" & PADC("DO QTY", 8) & "|" & PADC("DRAFT", 12) & "|" & PADC("DRAFT", 10) & "|" & PADC("DRAFT", 15)
invrep.Sections("Section1").Controls("Label22").Caption = PADC("TRANSPORT", 10) & "|" & PADC("NUMBER", 11) & "|" & PADC("DATE", 8) & "|" & PADC("", 8) & "|" & PADC("NUMBER", 12) & "|" & PADC("DATE", 10) & "|" & PADC("AMOUNT", 15)
invrep.Sections("Section1").Controls("Label23").Caption = String(80, "-")

dd1amt = rs2.Fields("draft_amt1")
dd2amt = rs2.Fields("draft_amt2")
dd3amt = rs2.Fields("draft_amt3")
othamt = 0

invrep.Sections("Section1").Controls("Label24").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no1"), 12) & "|" & PADC(rs2.Fields("draft_dt1"), 10) & "|" & padl(Format(dd1amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label25").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no2"), 12) & "|" & PADC(rs2.Fields("draft_dt2"), 10) & "|" & padl(Format(dd2amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label26").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no3"), 12) & "|" & PADC(rs2.Fields("draft_dt3"), 10) & "|" & padl(Format(dd3amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label27").Caption = PADR("OTHER QTY/AMOUNT:", 18) & PADR("", 20) & " Other amount deposited : " & padl(Format(othamt, "0######0.#0"), 15)
invrep.Sections("Section1").Controls("Label28").Caption = String(80, "-")

invrep.Sections("Section1").Controls("Label29").Caption = PADR("", 38) & "  Total amount deposited : " & padl(Format(dd1amt + dd2amt + dd3amt + othamt, "0######0.#0"), 15)
invrep.Sections("Section1").Controls("Label30").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label31").Caption = "DO.No.: " & PADR(dono, 11) & " Date: " & rs2.Fields("do_date") & " Grade: " & rs2.Fields("grade") & " DO Qty : " & rs2.Fields("do_qty")
invrep.Sections("Section1").Controls("Label32").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label33").Caption = PADC("Loading Date", 15) & PADC("Challan No", 15) & PADC("Truck No", 15) & PADC("Grade", 15) & PADC("Qty Lifted", 15)
invrep.Sections("Section1").Controls("Label34").Caption = String(80, "-")

lifted = CDbl(ntw.Caption) / 1000
invrep.Sections("Section1").Controls("Label35").Caption = PADC(odate, 15) & PADC(challanno.Caption, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
invrep.Sections("Section1").Controls("Label36").Caption = String(80, "-")

invrep.Sections("Section1").Controls("Label37").Caption = PADR("Tot.Lifted Qty:", 15) + PADR(LTrim(lqty.Caption), 10) & PADR("BILLED HEAD", 30) & padl("RATE", 12) & padl("VALUE", 13)
invrep.Sections("Section1").Controls("Label38").Caption = String(25, " ") & PADR("", 30) & String(25, "-")
invrep.Sections("Section1").Controls("Label39").Caption = String(25, " ") & PADR("Basic Value", 30) & padl(Format(CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 12) & padl(Format(lifted * CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 13)
excdata(8) = rs2.Fields("basic_rate")
excdata(9) = lifted * CDbl(rs2.Fields("basic_rate"))
invrep.Sections("Section1").Controls("Label40").Caption = String(25, " ") & PADR("Add on Price (WRC)", 30) & padl(CDbl(rs2.Fields("wrc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("wrc")), "#####0.#0"), 13)
excdata(10) = rs2.Fields("wrc")
excdata(11) = lifted * CDbl(rs2.Fields("wrc"))
invrep.Sections("Section1").Controls("Label41").Caption = String(25, " ") & PADR("Selective Loading Charge", 30) & padl(CDbl(rs2.Fields("slc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("slc")), "#####0.#0"), 13)
excdata(12) = rs2.Fields("slc")
excdata(13) = lifted * CDbl(rs2.Fields("slc"))
invrep.Sections("Section1").Controls("Label42").Caption = PADR("Item Code: 27011910", 25) & PADR("Weighment Charge", 30) & padl(CDbl(rs2.Fields("weighment_chg")), 12) & padl(Format(lifted * CDbl(rs2.Fields("weighment_chg")), "#####0.#0"), 13)
excdata(14) = rs2.Fields("weighment_chg")
excdata(15) = lifted * CDbl(rs2.Fields("weighment_chg"))
excdata(16) = "0.00"
excdata(17) = "0.00"
excgross = (lifted * CDbl(rs2.Fields("basic_rate"))) + (lifted * CDbl(rs2.Fields("wrc"))) + (lifted * CDbl(rs2.Fields("slc"))) + (lifted * CDbl(rs2.Fields("weighment_chg"))) + (lifted * CDbl(rs2.Fields("royalty"))) + (lifted * CDbl(rs2.Fields("sed")))
excdata(18) = excgross
invrep.Sections("Section1").Controls("Label43").Caption = PADR("Description of Goods", 30) & PADR("Excisable Gross", 32) & padl(Format(excgross, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label44").Caption = PADR("Raw Coal", 30) & PADR("Cent Excise @" + rs2.Fields("cent_exc_rate") + "%", 32) & padl(Format(excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label45").Caption = PADR("Destination: " & Trim(cadr), 30) & PADR("Ed Cess on Excise @" + rs2.Fields("edu_cess_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label46").Caption = PADR(" ", 30) & PADR("S/H Cess on Excise @" + rs2.Fields("high_edu_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label47").Caption = PADR("Type of Coal:", 25) & PADR("Royalty", 30) & PADR(CDbl(rs2.Fields("royalty")), 12) & padl(Format(lifted * CDbl(rs2.Fields("royalty")), "#####0.#0"), 13)
excdata(26) = rs2.Fields("royalty")
excdata(27) = lifted * CDbl(rs2.Fields("royalty"))
invrep.Sections("Section1").Controls("Label48").Caption = PADR("Non-Coking Coal", 25) & PADR("Stowing Excise Duty(SED)", 30) & padl(CDbl(rs2.Fields("sed")), 12) & padl(Format(lifted * CDbl(rs2.Fields("sed")), "#####0.#0"), 13)
excdata(28) = rs2.Fields("sed")
excdata(29) = lifted * CDbl(rs2.Fields("sed"))
excduty = (excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01)
totgross = excgross + excduty
excdata(19) = rs2.Fields("cent_exc_rate")
excdata(20) = excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01
excdata(21) = rs2.Fields("edu_cess_rate")
excdata(22) = excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01
excdata(23) = rs2.Fields("high_edu_rate")
excdata(24) = excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01
excdata(25) = excduty
invrep.Sections("Section1").Controls("Label49").Caption = PADR("", 25) & PADR("Total Excise Duty", 42) & padl(Format(excduty, "#####0.#0"), 13)
decpart = Right(Format(excduty, "######.#0"), 2)
excdut1 = Left(Format(excduty, "0#####.#0"), Len(Format(excduty, "0#####.#0")) - 3)
invrep.Sections("Section1").Controls("Label50").Caption = "Rupees " & worda(Val(excdut1)) & " and " & worda(CLng(decpart)) & " paise only"
invrep.Sections("Section1").Controls("Label51").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label52").Caption = String(25, " ") & PADR("Clean Energy Cess", 30) & padl(CDbl(rs2.Fields("clean_engy_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("clean_engy_cess")), "#####0.#0"), 13)
excdata(30) = rs2.Fields("clean_engy_cess")
excdata(31) = lifted * CDbl(rs2.Fields("clean_engy_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("clean_engy_cess"))
invrep.Sections("Section1").Controls("Label53").Caption = String(25, " ") & PADR("Road/ RE Cess", 30) & padl(CDbl(rs2.Fields("road_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("road_cess")), "#####0.#0"), 13)
excdata(32) = rs2.Fields("road_cess")
excdata(33) = lifted * CDbl(rs2.Fields("road_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("road_cess"))
invrep.Sections("Section1").Controls("Label54").Caption = String(25, " ") & PADR("PWD Cess/ MADA(Bazaar)", 30) & padl(CDbl(rs2.Fields("bazar_fee")), 12) & padl(Format(lifted * CDbl(rs2.Fields("bazar_fee")), "#####0.#0"), 13)
excdata(34) = rs2.Fields("bazar_fee")
excdata(35) = lifted * CDbl(rs2.Fields("bazar_fee"))
totgross = totgross + lifted * CDbl(rs2.Fields("bazar_fee"))
invrep.Sections("Section1").Controls("Label55").Caption = String(25, " ") & PADR("AMBH Cess", 30) & padl(CDbl(rs2.Fields("ambh_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("ambh_cess")), "#####0.#0"), 13)
excdata(36) = rs2.Fields("ambh_cess")
excdata(37) = lifted * CDbl(rs2.Fields("ambh_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("ambh_cess"))
invrep.Sections("Section1").Controls("Label56").Caption = String(25, " ") & PADR("Other Charges", 30) & padl(" ", 12) & padl(CDbl(rs2.Fields("other_charges")), 13)

excdata(38) = rs2.Fields("other_charges")
excdata(39) = lifted * CDbl(rs2.Fields("other_charges"))
MsgBox "1"
totgross = totgross + lifted * CDbl(rs2.Fields("other_charges"))
MsgBox "2"
excdata(40) = totgross
MsgBox "4"
invrep.Sections("Section1").Controls("Label57").Caption = String(25, " ") & String(55, "-")
MsgBox "5"
invrep.Sections("Section1").Controls("Label58").Caption = String(25, " ") & PADR("TOTAL VALUE", 40) & padl(Format(totgross, "#####0.#0"), 13)
MsgBox "6"
invrep.Sections("Section1").Controls("Label59").Caption = String(25, " ") & PADR("VAT/CST@ " + rs2.Fields("tax_percent") + "%", 30) & padl("JST", 12) & padl(Format(totgross * 0.01 * CDbl(rs2.Fields("tax_percent")), "#####0.#0"), 13)
MsgBox "4"

excdata(41) = rs2.Fields("tax_percent")
excdata(42) = totgross * 0.01 * CDbl(rs2.Fields("tax_percent"))
If IsNull(CDbl(rs2.Fields("tcs"))) Then rs2.Fields("tcs") = 0
tcs = totgross * 0.01 * CDbl(rs2.Fields("tcs"))
tcs = tcs + (excdata(42) * 0.01 * CDbl(rs2.Fields("tcs")))

totgross = totgross + (totgross * 0.01 * CDbl(rs2.Fields("tax_percent")))
excdata(43) = totgross
invrep.Sections("Section1").Controls("Label60").Caption = String(25, " ") & PADR("TCS@ " & rs2.Fields("tcs"), 30) & padl(" ", 12) & padl(Format(tcs, "#####0.#0"), 13)
excdata(44) = "1"
excdata(45) = tcs
invrep.Sections("Section1").Controls("Label61").Caption = String(25, " ") & String(55, "-")
totgross = totgross + tcs
excdata(46) = totgross
excdata(47) = "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6)
excdata(48) = odate.Caption
excdata(49) = rtype.Caption



invrep.Sections("Section1").Controls("Label62").Caption = String(25, " ") & PADR("Gross Value", 40) & padl(Format(totgross, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label63").Caption = String(25, " ") & String(55, "-")
invrep.Sections("Section1").Controls("Label64").Caption = "Gross value in words:"
decpart = Right(Format(totgross, "######.#0"), 2)
On Error Resume Next
totgross = Left(Format(totgross, "######.#0"), Len(Format(totgross, "######.#0")) - 3)
invrep.Sections("Section1").Controls("Label65").Caption = "Rupees " & worda(Val(totgross)) & "and " & worda(CLng(decpart)) & "paise only"
invrep.Sections("Section1").Controls("Label66").Caption = String(80, "-")
totgross = 0

invrep.Sections("Section1").Controls("Label67").Caption = PADR("Prepared By", 26) & PADR("Checked By", 26) & PADR("Auth. Signatory", 26)
invrep.Sections("Section1").Controls("Label68").Caption = PADR(" ", 45) & PADR("Under Jurisdiction of Jharkhand", 35)

rs2.Close
rs4.Close
If printop = 1 Then
invrep.Show vbModal
End If

End Sub


Private Sub invoicerep1()
Dim excgross As Double
Dim excduty As Double
Dim excdut1 As Long
Dim totgross As Double
Dim decpart As Double
Dim lifted As Double
Dim dd1amt As Double
Dim dd2amt As Double
Dim dd3amt As Double
Dim othamt As Double
Dim tcs As Double
Dim dsno As Long
Dim billfnd As Boolean

excdata(0) = "ROAD"
excdata(1) = "0"
excdata(2) = "0"
excdata(3) = "0"
excdata(4) = "27011910"
excdata(5) = "RAW COAL"
excdata(7) = "NC"

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from cadata where do_no='" & Trim(dono.Caption) & "'", co, adOpenKeyset, adLockOptimistic
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
Else
    MsgBox "Relevant CA Data not found"
    Exit Sub
End If

Set rs4 = New ADODB.Recordset
rs4.Open "Select * from excisedata", co, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
    rs4.MoveFirst
Else
    MsgBox "Relevant Excise Data not found"
    Exit Sub
End If

excgross = 0
excduty = 0
totgross = 0

Set rs5 = New ADODB.Recordset
rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' and sl_no = '" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs5.RecordCount = 0 Then
    rs5.Close
    Set rs5 = New ADODB.Recordset
    rs5.Open "Select * from billdata where season='" & Trim(sesson.Caption) & "' order by billno", co, adOpenKeyset, adLockOptimistic
    If rs5.RecordCount > 0 Then
        rs5.MoveLast
        dsno = rs5.Fields("billno").Value + 1
    Else
        dsno = 1
    End If

    rs5.AddNew
    rs5.Fields(0).Value = Trim(sesson.Caption)
    rs5.Fields(1).Value = dsno
    rs5.Fields(2).Value = Trim(Text1.Text)
    rs5.Update
Else
    dsno = rs5.Fields("billno").Value
End If
rs5.Close

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set invrep.DataSource = rs1
invrep.Sections("Section1").Controls("Text1").DataField = "SEASON"
End If

invrep.Sections("Section1").Controls("Label1").Caption = "Original/ Duplicate/ Triplicate/ Quadruplicate"
invrep.Sections("Section1").Controls("Label2").Caption = "for Road"

invrep.Sections("Section1").Controls("Label3").Caption = PADR("           BHARAT COKING COAL LIMITED             ", 50) & PADR("Pre Authenticated", 30)
invrep.Sections("Section1").Controls("Label4").Caption = PADR("       A Subsidiary of Coal India Limited         ", 50)
invrep.Sections("Section1").Controls("Label5").Caption = PADR("       Regd. Office: Koyla Bhavan, DHANBAD        ", 50) & PADR("for Bharat Coking Coal Limited", 30)
invrep.Sections("Section1").Controls("Label6").Caption = String(50, " ") & PADR("Authorized Signatory", 30)
invrep.Sections("Section1").Controls("Label7").Caption = PADC("Tax Invoice under rule 11 of C/E Rules 2002", 80)
invrep.Sections("Section1").Controls("Label8").Caption = "Type of Customer: " & rtype.Caption
invrep.Sections("Section1").Controls("Label9").Caption = "Bill No: " & "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6) & "       Bill Date: " & odate.Caption
invrep.Sections("Section1").Controls("Label10").Caption = "TAG No : " & Left(rs1.Fields("tag").Value, 4) & "XXXXXXXXXX" & Right(rs1.Fields("tag").Value, 4)

invrep.Sections("Section1").Controls("Label11").Caption = PADR("Unit Name: ", 16) & PADR(collname, 23) & PADR(" Purchaser Name: ", 17) & PADR(cname, 24)
invrep.Sections("Section1").Controls("Label12").Caption = PADR("Address: ", 16) & PADR(Trim(main.Label3.Caption), 23) & PADR(" Address: ", 17) & PADR(rs2.Fields("destination"), 24)
excdata(6) = rs2.Fields("destination")
invrep.Sections("Section1").Controls("Label13").Caption = PADR("Sel.Exc.Reg.No.: ", 16) & PADR(rs4.Fields("u_excno"), 23) & PADR(" Pur.Exc.RegNo.: ", 17) & PADR(rs2.Fields("exc_reg_no"), 24)
invrep.Sections("Section1").Controls("Label14").Caption = PADR("Range: ", 16) & PADR(rs4.Fields("u_range"), 23) & PADR(" Range: ", 17) & PADR(rs2.Fields("range"), 24)
invrep.Sections("Section1").Controls("Label15").Caption = PADR("Division: ", 16) & PADR(rs4.Fields("u_division"), 23) & PADR(" Division: ", 17) & PADR(rs2.Fields("division"), 24)
invrep.Sections("Section1").Controls("Label16").Caption = PADR("Commissionerate: ", 16) & PADR(rs4.Fields("u_commissionrate"), 23) & PADR(" Commissionerate: ", 17) & PADR(rs2.Fields("commissionerate"), 24)
invrep.Sections("Section1").Controls("Label17").Caption = PADR("Vat Tin No.: ", 16) & PADR(rs4.Fields("u_tinno"), 23) & PADR(" Vat Tin No.: ", 17) & PADR(rs2.Fields("vat_tin_no"), 24)
invrep.Sections("Section1").Controls("Label18").Caption = PADR("CST No.: ", 16) & PADR(rs4.Fields("u_cstno"), 23) & PADR(" CST No.: ", 17) & PADR(rs2.Fields("cst_no"), 24)
invrep.Sections("Section1").Controls("Label19").Caption = PADR("PAN No.: ", 16) & PADR(rs4.Fields("u_panno"), 23) & PADR(" PAN No.: ", 17) & PADR(rs2.Fields("PAN"), 24)
invrep.Sections("Section1").Controls("Label20").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label21").Caption = PADC("MODE OF", 10) & "|" & PADC("APPLICATION", 20) & "|" & PADC("DO QTY", 8) & "|" & PADC("DRAFT", 12) & "|" & PADC("DRAFT", 10) & "|" & PADC("DRAFT", 15)
invrep.Sections("Section1").Controls("Label22").Caption = PADC("TRANSPORT", 10) & "|" & PADC("NUMBER", 11) & "|" & PADC("DATE", 8) & "|" & PADC("", 8) & "|" & PADC("NUMBER", 12) & "|" & PADC("DATE", 10) & "|" & PADC("AMOUNT", 15)
invrep.Sections("Section1").Controls("Label23").Caption = String(80, "-")

dd1amt = rs2.Fields("draft_amt1")
dd2amt = rs2.Fields("draft_amt2")
dd3amt = rs2.Fields("draft_amt3")
othamt = 0

invrep.Sections("Section1").Controls("Label24").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no1"), 12) & "|" & PADC(rs2.Fields("draft_dt1"), 10) & "|" & padl(Format(dd1amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label25").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no2"), 12) & "|" & PADC(rs2.Fields("draft_dt2"), 10) & "|" & padl(Format(dd2amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label26").Caption = PADC("ROAD", 10) & "|" & PADC(rs2.Fields("appl_no"), 11) & "|" & PADC(rs2.Fields("appl_date"), 8) & "|" & PADC(rs2.Fields("do_qty"), 8) & "|" & PADC(rs2.Fields("draft_no3"), 12) & "|" & PADC(rs2.Fields("draft_dt3"), 10) & "|" & padl(Format(dd3amt, "0#######.#0"), 15)
invrep.Sections("Section1").Controls("Label27").Caption = PADR("OTHER QTY/AMOUNT:", 18) & PADR("", 20) & " Other amount deposited : " & padl(Format(othamt, "0######0.#0"), 15)
invrep.Sections("Section1").Controls("Label28").Caption = String(80, "-")

invrep.Sections("Section1").Controls("Label29").Caption = PADR("", 38) & "  Total amount deposited : " & padl(Format(dd1amt + dd2amt + dd3amt + othamt, "0######0.#0"), 15)
invrep.Sections("Section1").Controls("Label30").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label31").Caption = "DO.No.: " & PADR(dono, 11) & " Date: " & rs2.Fields("do_date") & " Grade: " & rs2.Fields("grade") & " DO Qty : " & rs2.Fields("do_qty")
invrep.Sections("Section1").Controls("Label32").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label33").Caption = PADC("Loading Date", 15) & PADC("Challan No", 15) & PADC("Truck No", 15) & PADC("Grade", 15) & PADC("Qty Lifted", 15)
invrep.Sections("Section1").Controls("Label34").Caption = String(80, "-")

lifted = CDbl(ntw.Caption) / 1000
invrep.Sections("Section1").Controls("Label35").Caption = PADC(odate, 15) & PADC(challanno.Caption, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
invrep.Sections("Section1").Controls("Label36").Caption = String(80, "-")

invrep.Sections("Section1").Controls("Label37").Caption = PADR("Tot.Lifted Qty:", 15) + PADR(LTrim(lqty.Caption), 10) & PADR("BILLED HEAD", 30) & padl("RATE", 12) & padl("VALUE", 13)
invrep.Sections("Section1").Controls("Label38").Caption = String(25, " ") & PADR("", 30) & String(25, "-")
invrep.Sections("Section1").Controls("Label39").Caption = String(25, " ") & PADR("Basic Value", 30) & padl(Format(CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 12) & padl(Format(lifted * CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 13)
excdata(8) = rs2.Fields("basic_rate")
excdata(9) = lifted * CDbl(rs2.Fields("basic_rate"))
invrep.Sections("Section1").Controls("Label40").Caption = String(25, " ") & PADR("Add on Price (WRC)", 30) & padl(CDbl(rs2.Fields("wrc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("wrc")), "#####0.#0"), 13)
excdata(10) = rs2.Fields("wrc")
excdata(11) = lifted * CDbl(rs2.Fields("wrc"))
invrep.Sections("Section1").Controls("Label41").Caption = String(25, " ") & PADR("Selective Loading Charge", 30) & padl(CDbl(rs2.Fields("slc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("slc")), "#####0.#0"), 13)
excdata(12) = rs2.Fields("slc")
excdata(13) = lifted * CDbl(rs2.Fields("slc"))
invrep.Sections("Section1").Controls("Label42").Caption = PADR("Item Code: 27011910", 25) & PADR("Weighment Charge", 30) & padl(CDbl(rs2.Fields("weighment_chg")), 12) & padl(Format(lifted * CDbl(rs2.Fields("weighment_chg")), "#####0.#0"), 13)
excdata(14) = rs2.Fields("weighment_chg")
excdata(15) = lifted * CDbl(rs2.Fields("weighment_chg"))
excdata(16) = "0.00"
excdata(17) = "0.00"
excgross = (lifted * CDbl(rs2.Fields("basic_rate"))) + (lifted * CDbl(rs2.Fields("wrc"))) + (lifted * CDbl(rs2.Fields("slc"))) + (lifted * CDbl(rs2.Fields("weighment_chg"))) + (lifted * CDbl(rs2.Fields("royalty"))) + (lifted * CDbl(rs2.Fields("sed")))
excdata(18) = excgross
invrep.Sections("Section1").Controls("Label43").Caption = PADR("Description of Goods", 30) & PADR("Excisable Gross", 32) & padl(Format(excgross, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label44").Caption = PADR("Raw Coal", 30) & PADR("Cent Excise @" + rs2.Fields("cent_exc_rate") + "%", 32) & padl(Format(excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label45").Caption = PADR("Destination: " & Trim(cadr), 30) & PADR("Ed Cess on Excise @" + rs2.Fields("edu_cess_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label46").Caption = PADR(" ", 30) & PADR("S/H Cess on Excise @" + rs2.Fields("high_edu_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label47").Caption = PADR("Type of Coal:", 25) & PADR("Royalty", 30) & PADR(CDbl(rs2.Fields("royalty")), 12) & padl(Format(lifted * CDbl(rs2.Fields("royalty")), "#####0.#0"), 13)
excdata(26) = rs2.Fields("royalty")
excdata(27) = lifted * CDbl(rs2.Fields("royalty"))
invrep.Sections("Section1").Controls("Label48").Caption = PADR("Non-Coking Coal", 25) & PADR("Stowing Excise Duty(SED)", 30) & padl(CDbl(rs2.Fields("sed")), 12) & padl(Format(lifted * CDbl(rs2.Fields("sed")), "#####0.#0"), 13)
excdata(28) = rs2.Fields("sed")
excdata(29) = lifted * CDbl(rs2.Fields("sed"))
excduty = (excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01) + (excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01)
totgross = excgross + excduty
excdata(19) = rs2.Fields("cent_exc_rate")
excdata(20) = excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01
excdata(21) = rs2.Fields("edu_cess_rate")
excdata(22) = excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01
excdata(23) = rs2.Fields("high_edu_rate")
excdata(24) = excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01
excdata(25) = excduty
invrep.Sections("Section1").Controls("Label49").Caption = PADR("", 25) & PADR("Total Excise Duty", 42) & padl(Format(excduty, "#####0.#0"), 13)
decpart = Right(Format(excduty, "######.#0"), 2)
excdut1 = Left(Format(excduty, "0#####.#0"), Len(Format(excduty, "0#####.#0")) - 3)
invrep.Sections("Section1").Controls("Label50").Caption = "Rupees " & worda(Val(excdut1)) & " and " & worda(CLng(decpart)) & " paise only"
invrep.Sections("Section1").Controls("Label51").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label52").Caption = String(25, " ") & PADR("Clean Energy Cess", 30) & padl(CDbl(rs2.Fields("clean_engy_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("clean_engy_cess")), "#####0.#0"), 13)
excdata(30) = rs2.Fields("clean_engy_cess")
excdata(31) = lifted * CDbl(rs2.Fields("clean_engy_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("clean_engy_cess"))
invrep.Sections("Section1").Controls("Label53").Caption = String(25, " ") & PADR("Road/ RE Cess", 30) & padl(CDbl(rs2.Fields("road_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("road_cess")), "#####0.#0"), 13)
excdata(32) = rs2.Fields("road_cess")
excdata(33) = lifted * CDbl(rs2.Fields("road_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("road_cess"))
invrep.Sections("Section1").Controls("Label54").Caption = String(25, " ") & PADR("PWD Cess/ MADA(Bazaar)", 30) & padl(CDbl(rs2.Fields("bazar_fee")), 12) & padl(Format(lifted * CDbl(rs2.Fields("bazar_fee")), "#####0.#0"), 13)
excdata(34) = rs2.Fields("bazar_fee")
excdata(35) = lifted * CDbl(rs2.Fields("bazar_fee"))
totgross = totgross + lifted * CDbl(rs2.Fields("bazar_fee"))
invrep.Sections("Section1").Controls("Label55").Caption = String(25, " ") & PADR("AMBH Cess", 30) & padl(CDbl(rs2.Fields("ambh_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("ambh_cess")), "#####0.#0"), 13)
excdata(36) = rs2.Fields("ambh_cess")
excdata(37) = lifted * CDbl(rs2.Fields("ambh_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("ambh_cess"))
invrep.Sections("Section1").Controls("Label56").Caption = String(25, " ") & PADR("Other Charges", 30) & padl(" ", 12) & padl(CDbl(rs2.Fields("other_charges")), 13)

excdata(38) = rs2.Fields("other_charges")
excdata(39) = lifted * CDbl(rs2.Fields("other_charges"))
totgross = totgross + lifted * CDbl(rs2.Fields("other_charges"))
excdata(40) = totgross
invrep.Sections("Section1").Controls("Label57").Caption = String(25, " ") & String(55, "-")
invrep.Sections("Section1").Controls("Label58").Caption = String(25, " ") & PADR("TOTAL VALUE", 40) & padl(Format(totgross, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label59").Caption = String(25, " ") & PADR("VAT/CST@ " + rs2.Fields("tax_percent") + "%", 30) & padl("JST", 12) & padl(Format(totgross * 0.01 * CDbl(rs2.Fields("tax_percent")), "#####0.#0"), 13)

excdata(41) = rs2.Fields("tax_percent")
excdata(42) = totgross * 0.01 * CDbl(rs2.Fields("tax_percent"))
If IsNull(CDbl(rs2.Fields("tcs"))) Then rs2.Fields("tcs") = 0
tcs = totgross * 0.01 * CDbl(rs2.Fields("tcs"))
tcs = tcs + (excdata(42) * 0.01 * CDbl(rs2.Fields("tcs")))

totgross = totgross + (totgross * 0.01 * CDbl(rs2.Fields("tax_percent")))
excdata(43) = totgross
invrep.Sections("Section1").Controls("Label60").Caption = String(25, " ") & PADR("TCS@ " & rs2.Fields("tcs"), 30) & padl(" ", 12) & padl(Format(tcs, "#####0.#0"), 13)
excdata(44) = "1"
excdata(45) = tcs
invrep.Sections("Section1").Controls("Label61").Caption = String(25, " ") & String(55, "-")
totgross = totgross + tcs
excdata(46) = totgross
excdata(47) = "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6)
excdata(48) = odate.Caption
excdata(49) = rtype.Caption



invrep.Sections("Section1").Controls("Label62").Caption = String(25, " ") & PADR("Gross Value", 40) & padl(Format(totgross, "#####0.#0"), 13)
invrep.Sections("Section1").Controls("Label63").Caption = String(25, " ") & String(55, "-")
invrep.Sections("Section1").Controls("Label64").Caption = "Gross value in words:"
decpart = Right(Format(totgross, "######.#0"), 2)
On Error Resume Next
totgross = Left(Format(totgross, "######.#0"), Len(Format(totgross, "######.#0")) - 3)
invrep.Sections("Section1").Controls("Label65").Caption = "Rupees " & worda(Val(totgross)) & "and " & worda(CLng(decpart)) & "paise only"
invrep.Sections("Section1").Controls("Label66").Caption = String(80, "-")
totgross = 0

invrep.Sections("Section1").Controls("Label67").Caption = PADR("Prepared By", 26) & PADR("Checked By", 26) & PADR("Auth. Signatory", 26)
invrep.Sections("Section1").Controls("Label68").Caption = PADR(" ", 45) & PADR("Under Jurisdiction of Jharkhand", 35)

rs2.Close
rs4.Close
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
unlocktxt
End Sub

Private Sub Command6_Click()
printop = 0
printrep
printop = 0

If Val(Text2.Text) > 0 Then
invoicerep1
'invoicerep
a = writefile
'writeora = False
'writeserver
'If writeora = False Then
'    MsgBox "Data transfer to oracle server failed"
'Else
'    Set rs1 = New ADODB.Recordset
'    rs1.Open "Select * from special where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
'    If rs1.RecordCount > 0 Then
'        rs1.MoveFirst
'        rs1.Fields("data38") = "1"
'        rs1.Update
'    End If
'    rs1.Close
'    MsgBox "Data transfered to oracle server successfully"
'End If
End If
End Sub



Private Sub Form_Load()
Me.Picture = main.Picture

Conn
username = loginname
unlocktxt
DTPicker1.Value = Date

FTP1.HostName = ips(2)
FTP1.username = unames(2)
FTP1.PassWord = passws(2)
FTP1.ChangeRemoteDir (paths(2))
printop = 1
End Sub

Private Sub unlocktxt()
op2.Caption = ""
sesson.Caption = ""
vno.Caption = ""
cname.Caption = ""
cadr.Caption = ""
op1.Caption = ""
intime.Caption = ""
chno.Caption = ""
material.Caption = ""
fwt.Caption = ""
ntw.Caption = ""
Text2.Text = Val(0)

op2.Caption = ""
sesson.Caption = ""
vno.Caption = ""
cco.Caption = ""
cname.Caption = ""
cadr.Caption = ""
op1.Caption = ""

intime.Caption = ""
chno.Caption = ""
material.Caption = ""
fwt.Caption = ""
rlw.Caption = ""

collcode.Caption = ""
collname.Caption = ""
challanno.Caption = ""
challandate.Caption = ""
inshift.Caption = ""
outshift.Caption = ""
dono.Caption = ""
rtype.Caption = ""
''DTPicker2.Value = ""
'DTPicker3.Value = ""
oqty.Caption = ""
lqty.Caption = ""
bqty.Caption = ""
Text2.Text = ""
ntw.Caption = ""
Text1.Text = ""
'Text1.SetFocus
End Sub

Private Sub datashow()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
    op2.Caption = username
    sesson.Caption = rs1.Fields("season").Value
    vno.Caption = rs1.Fields("V_no").Value
    
    cco.Caption = rs1.Fields("custcd").Value
    cname.Caption = rs1.Fields("purchaser").Value
    cadr.Caption = rs1.Fields("dest").Value
    op1.Caption = rs1.Fields("o_name").Value
    
    If Not IsNull(rs1.Fields("o2_name").Value) Then op2.Caption = rs1.Fields("o2_name").Value
    
    DTPicker1.Value = rs1.Fields("date_in").Value
    intime.Caption = rs1.Fields("time_in").Value
    chno.Caption = rs1.Fields("tm_code").Value
'    material.Caption = rs1.Fields("m_name").Value
    fwt.Caption = rs1.Fields("first_wt").Value
    rlw.Caption = rs1.Fields("RLW").Value
        
    If Not IsNull(rs1.Fields("date_out").Value) Then
        odate.Caption = rs1.Fields("date_out").Value
        outtime.Caption = rs1.Fields("time_out").Value
    End If
    collcode.Caption = rs1.Fields("coll_code").Value
    collname.Caption = rs1.Fields("coll_desc").Value
    If Len(rs1.Fields("challan_no").Value) > 0 Then
    challanno.Caption = rs1.Fields("challan_no").Value
    challandate.Caption = rs1.Fields("challan_date").Value
    DTPicker4.Value = rs1.Fields("challan_date").Value
    End If
    inshift.Caption = rs1.Fields("shift_in").Value
    outshift.Caption = Trim(main.Label6)
    dono.Caption = rs1.Fields("do_no").Value
    rtype.Caption = "FSA/MOU/E-AUC/OTHER"

DTPicker2.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
DTPicker2.Year = Mid(rs1.Fields("do_start_date").Value, 1, 4)
DTPicker2.Month = Mid(rs1.Fields("do_start_date").Value, 5, 2)
DTPicker2.Day = Mid(rs1.Fields("do_start_date").Value, 7, 2)

DTPicker3.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
DTPicker3.Year = Mid(rs1.Fields("do_end_date").Value, 1, 4)
DTPicker3.Month = Mid(rs1.Fields("do_end_date").Value, 5, 2)
DTPicker3.Day = Mid(rs1.Fields("do_end_date").Value, 7, 2)

oqty.Caption = Val(rs1.Fields("do_qty").Value) * 1000
    
    
    
Set rs3 = New ADODB.Recordset
rs3.Open "Select sum(Abs(second_wt-first_wt)) from special where do_no='" & Trim(dono.Caption) & "' and second_wt>0", co, adOpenKeyset, adLockOptimistic

If rs3.RecordCount > 0 Then
If IsNumeric(rs3.Fields(0)) Then
    lqty.Caption = Val(rs3.Fields(0))
    bqty.Caption = Val(oqty.Caption) - Val(rs3.Fields(0))
    balqty = bqty.Caption
    lifqty = lqty.Caption
Else
    lqty.Caption = 0
    bqty.Caption = Val(oqty.Caption)
    balqty = bqty.Caption
    lifqty = lqty.Caption
End If
End If
    
    Text2.Text = rs1.Fields("second_wt").Value
    ntw.Caption = Abs(rs1.Fields("second_wt").Value - rs1.Fields("first_wt").Value)
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
'KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
        datashow
        Text1.Enabled = False
    End If
End If
End Sub

Private Sub Text2_Change()
ntw.Caption = Abs(Val(fwt.Caption) - Val(Text2.Text))
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text2.Text) = 0 Then
Command3.SetFocus
End If
End If

End Sub

