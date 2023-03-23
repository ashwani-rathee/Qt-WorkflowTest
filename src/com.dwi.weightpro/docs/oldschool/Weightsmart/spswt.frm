VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form spswt 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   315
      Left            =   7920
      TabIndex        =   82
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   43048
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5160
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   6960
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   6900
   End
   Begin VB.ComboBox readtags 
      Height          =   315
      Left            =   1440
      TabIndex        =   77
      Text            =   "Combo1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2040
      Top             =   6900
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   6900
   End
   Begin MSComCtl2.DTPicker tmpdt 
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67108865
      CurrentDate     =   41165
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1560
      Top             =   6900
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   14895
      Begin VB.TextBox Text18 
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
         TabIndex        =   83
         Top             =   4080
         Width           =   1695
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   120
         Width           =   5415
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
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
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Unload Form"
         Top             =   60
         Visible         =   0   'False
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   9600
         TabIndex        =   72
         Top             =   5820
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   11160
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   71
         Top             =   3120
         Width           =   3615
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   7440
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   70
         Top             =   3120
         Width           =   3615
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   11160
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   69
         Top             =   480
         Width           =   3615
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   7440
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   68
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         Height          =   1035
         Left            =   10800
         TabIndex        =   65
         Top             =   5700
         Width           =   4095
         Begin VB.CommandButton Command1 
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
            Height          =   735
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   180
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "TRANSACTION CANCEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   180
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   2415
         Left            =   4080
         TabIndex        =   37
         Top             =   2760
         Width           =   3255
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1560
            TabIndex        =   43
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   67108865
            CurrentDate     =   39961
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1560
            TabIndex        =   44
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   67108865
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
            TabIndex        =   56
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
            Left            =   1560
            TabIndex        =   55
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
            TabIndex        =   54
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
            Left            =   1560
            TabIndex        =   53
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
            TabIndex        =   52
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
            Left            =   1560
            TabIndex        =   51
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
            TabIndex        =   45
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            Left            =   1560
            TabIndex        =   40
            Top             =   480
            Width           =   2925
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
            TabIndex        =   39
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
            Left            =   1560
            TabIndex        =   38
            Top             =   240
            Width           =   2085
         End
      End
      Begin VB.CommandButton Command6 
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
         Left            =   3720
         TabIndex        =   1
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   23
         Top             =   5640
         Width           =   1575
      End
      Begin VB.TextBox Text1 
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
         Height          =   330
         Left            =   2040
         TabIndex        =   0
         Top             =   600
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5400
         TabIndex        =   4
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
         Format          =   67108865
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   2040
         TabIndex        =   62
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   67108865
         CurrentDate     =   40651
      End
      Begin VB.Shape door1 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5880
         Top             =   5580
         Width           =   375
      End
      Begin VB.Shape door2 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   6840
         Top             =   5580
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
         Left            =   5880
         TabIndex        =   79
         Top             =   5280
         Width           =   375
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
         Left            =   6840
         TabIndex        =   78
         Top             =   5280
         Width           =   375
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
         Height          =   555
         Left            =   4800
         TabIndex        =   76
         Top             =   6120
         Width           =   5775
      End
      Begin VB.Label Label12 
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
         Left            =   7440
         TabIndex        =   75
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label13 
         Height          =   315
         Left            =   2160
         TabIndex        =   73
         Top             =   6360
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Hologram No"
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
         TabIndex        =   63
         Top             =   4200
         Width           =   1395
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
         Left            =   5400
         TabIndex        =   61
         Top             =   1800
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
         Left            =   4200
         TabIndex        =   60
         Top             =   1800
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
         Left            =   5400
         TabIndex        =   59
         Top             =   1560
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
         Left            =   4200
         TabIndex        =   58
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   57
         Top             =   4560
         Visible         =   0   'False
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
         TabIndex        =   50
         Top             =   3600
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
         TabIndex        =   49
         Top             =   3600
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
         TabIndex        =   48
         Top             =   3360
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
         TabIndex        =   47
         Top             =   3360
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
         TabIndex        =   46
         Top             =   1440
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
         TabIndex        =   36
         Top             =   3120
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
         TabIndex        =   35
         Top             =   3120
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
         TabIndex        =   34
         Top             =   1440
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
         Left            =   4200
         TabIndex        =   33
         Top             =   1200
         Width           =   975
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
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label18 
         Caption         =   "O Date"
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
         Left            =   4200
         TabIndex        =   31
         Top             =   600
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5400
         TabIndex        =   30
         Top             =   600
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   960
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
         Left            =   4200
         TabIndex        =   28
         Top             =   960
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         Left            =   2160
         TabIndex        =   25
         Top             =   6120
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
         TabIndex        =   24
         Top             =   6120
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
         TabIndex        =   22
         Top             =   5700
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
         Left            =   5640
         TabIndex        =   21
         Top             =   2400
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
         Left            =   5640
         TabIndex        =   20
         Top             =   2160
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
         Left            =   4200
         TabIndex        =   19
         Top             =   2400
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
         Left            =   4200
         TabIndex        =   18
         Top             =   2160
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
         Left            =   2160
         TabIndex        =   17
         Top             =   5280
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
         TabIndex        =   16
         Top             =   2880
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
         TabIndex        =   15
         Top             =   2640
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
         TabIndex        =   14
         Top             =   2400
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
         TabIndex        =   13
         Top             =   1920
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
         TabIndex        =   12
         Top             =   1200
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
         TabIndex        =   11
         Top             =   5280
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
         TabIndex        =   10
         Top             =   2880
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
         TabIndex        =   9
         Top             =   2640
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
         TabIndex        =   8
         Top             =   2400
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
         TabIndex        =   7
         Top             =   1920
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
         TabIndex        =   6
         Top             =   1200
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
         TabIndex        =   5
         Top             =   600
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
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP FTP2 
      Left            =   6240
      OleObjectBlob   =   "spswt.frx":0000
      Top             =   6900
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP FTP1 
      Left            =   1080
      OleObjectBlob   =   "spswt.frx":0024
      Top             =   7020
   End
End
Attribute VB_Name = "spswt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim doorin As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim tctr As Integer
Dim tctr1 As Integer
Dim wtmode As String
Dim tagtrips As Integer
Dim tripsdone As Integer
Dim writeora As Boolean

Private WithEvents cFTP As clsFTP
Attribute cFTP.VB_VarHelpID = -1
Dim balqty As String
Dim lifqty As String
Dim flag As Integer
Dim excdata(50) As String
Dim printop As Integer

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

Private Sub dataset()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from special where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
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

writestr = padl(Trim(rtype), 1) & "|" & padl(Trim(dono), 11) & "|" & padl(trimdate(DTPicker2.Value), 8) & "|" & padl(trimdate(DTPicker3.Value), 8) & "|" & padl(Trim(cco), 6) & "|" & padl(Trim(collcode), 4) & "|" & padl(Trim(chno), 10) & "|" & padl(Format(weightord, "0####.##0"), 9) & "|" & padl(Format(Trim(CStr(weightdiff)), "0####.##0"), 9) & "|" & padl(Format(weightbal, "0####.##0"), 9) & "|" & padl(Trim(Text1), 12) & "|" & padl(Format(CDbl(rlw) / 1000, "0####.##0"), 9) & "|" & padl(Trim(vno), 15) & "|" & padl(Trim(Text18.Text), 8) & "|" & padl(trimdate(DTPicker4.Value), 8) & "|" & padl(trimdate(DTPicker1.Value), 8) & "|" & padl(Trim(ftime1), 6) & "|" & padl(trimdate(odate), 8) & "|" & padl(Trim(ftime2), 6) & "|" & padl(Format(CDbl(Text2) / 1000, "0#####.##0"), 10) & "|" & padl(Format(CDbl(fwt) / 1000, "0#####.##0"), 10) & "|" & padl(Format(CDbl(ntw) / 1000, "0#####.##0"), 10) & "|" & padl(Trim(cadr), 2) & "|" & padl(Trim(main.Label6), 1) & "|" & padl(Trim(op2), 6) & "|" _
 + padl(excdata(0), 4) + "|" + padl(excdata(1), 10) + "|" + padl(excdata(2), 13) + "|" + padl(excdata(3), 13) + "|" + padl(excdata(4), 8) + "|" + padl(excdata(5), 8) + "|" + padl(excdata(6), 20) + "|" + padl(excdata(7), 2) + "|" + padl(excdata(8), 8) + "|" + padl(excdata(9), 13) + "|" + padl(excdata(10), 8) + "|" + padl(excdata(11), 13) + "|" + padl(excdata(12), 8) + "|" + padl(excdata(13), 13) + "|" + padl(excdata(14), 8) + "|" + padl(excdata(15), 13) + "|" + padl(excdata(16), 8) + "|" + padl(excdata(17), 13) + "|" + padl(excdata(18), 13) + "|" + padl(excdata(19), 8) + "|" + padl(excdata(20), 13) + "|" + padl(excdata(21), 8) + "|" + padl(excdata(22), 13) + "|" + padl(excdata(23), 8) + "|" + padl(excdata(24), 13) + "|" _
 + padl(excdata(25), 13) + "|" + padl(excdata(26), 8) + "|" + padl(excdata(27), 13) + "|" + padl(excdata(28), 8) + "|" + padl(excdata(29), 13) + "|" + padl(excdata(30), 8) + "|" + padl(excdata(31), 13) + "|" + padl(excdata(32), 8) + "|" + padl(excdata(33), 13) + "|" + padl(excdata(34), 8) + "|" + padl(excdata(35), 13) + "|" + padl(excdata(36), 8) + "|" + padl(excdata(37), 13) + "|" + padl(excdata(38), 8) + "|" + padl(excdata(39), 13) + "|" + padl(excdata(40), 13) + "|" + padl(excdata(41), 8) + "|" + padl(excdata(42), 13) + "|" + padl(excdata(43), 13) + "|" + padl(excdata(44), 8) + "|" + padl(excdata(45), 13) + "|" + padl(excdata(46), 13) + "|" + padl(excdata(47), 32) + "|" + padl(excdata(48), 12) + "|" + padl(excdata(49), 2) + "|" + vbCrLf

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
rs6.Fields("challan_no") = Text18.Text
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




Private Sub bqty_Change()
If Val(bqty.Caption) < 0 Then
    MsgBox "Value of net quantity exceeds DO Value by " + bqty.Caption + " Kg, Weighment Terminated"
    Command1.Enabled = False
End If
End Sub


Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
If timecheck(cco.Caption) = False Then
    MsgBox "Weighment not allowed..Invalid Time Range"
    Command6.Enabled = False
    Command1.Enabled = False
End If

If Trim(cco.Caption) = "" Then
    MsgBox "Please select valid first weighment", vbInformation, "Balance overflow"
    Command1.Enabled = False
    Exit Sub
End If

If Val(bqty.Caption) < 0 Then
    MsgBox "Net weight exceeds balance quantity, weighment not possible", vbInformation, "Balance overflow "
    Command1.Enabled = False
    Exit Sub
End If

If Val(Text2.Text) < 20 Then
    MsgBox "Weight is too Low  ", vbInformation, "Weight is too low "
    Command1.Enabled = False
    Exit Sub
End If

If Val(ntw.Caption) < chkwt2 And ntw <> 0 Then
    MsgBox "Weight is less than minimum weight", vbInformation, "Weight is too low"
    Command1.Enabled = False
    Exit Sub
End If


If Val(Text2) > Val(rlw.Caption) Then
    MsgBox "Net weight exceeds RLW, please check weight", vbInformation, "RLW Violation"
    Command1.Enabled = False
    Exit Sub
End If

If Trim(Text1.Text) = "" Then
    MsgBox "Serial No does not Exist", vbCritical, "Error"
    Command1.Enabled = False
    Exit Sub
Else
    Set rs1 = New ADODB.Recordset
    rs1.Open "Select * from special where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs1.RecordCount > 0 Then
        rs1.MoveFirst
    Else
        MsgBox "wrong serial"
        Text1.SetFocus
        Command1.Enabled = False
    Exit Sub
    End If
End If

a = vbNo
a = MsgBox("Want to Update data", vbYesNo, "Update?")
If a = vbYes Then


tmpdt.Value = Date

On Error GoTo errm1
Conn2
If rfidcode = 1 Then
    Conn1
    co1.Execute "update special set o2_name='" + op2.Caption + "', time_out='" + outtime.Caption _
    + "', date_out='" + Format(Year(tmpdt.Value), "00##") + Format(Month(tmpdt.Value), "0#") + Format(Day(tmpdt.Value), "0#") _
    + "',second_wt=" + Trim(Text2.Text) + ", balance_qty=" + Trim(bqty.Caption) _
    + ", challan_date='" + Format(Year(tmpdt.Value), "00##") + Format(Month(tmpdt.Value), "0#") + Format(Day(tmpdt.Value), "0#") _
    + "', challan_no='" + Trim(Text18.Text) + "' where sl_no='" + Trim(Text1.Text) + "' and tag= '" + Trim(Texttag.Text) + "'"
    
    On Error GoTo errm
    If tripsdone >= tagtrips Then
        co1.Execute "update tags set valid='0',trips_done=" + CStr(tripsdone) + " where tagno = '" + Trim(Texttag.Text) + "' and tsno='" + Trim(Text1.Text) + "'"
    Else
        co1.Execute "update tags set valid='1',trips_done=" + CStr(tripsdone) + " where tagno = '" + Trim(Texttag.Text) + "' and tsno='" + Trim(Text1.Text) + "'"
    End If
Else
    Set rs5 = New ADODB.Recordset
    rs5.Open "Select * from  specialtmp where date_in = #" & Format(CDate(DTPicker1.Value), "mm/dd/yyyy") & "# and sl_no='" & Trim(Text1.Text) & "'", co2, adOpenKeyset, adLockOptimistic
    If rs5.RecordCount = 0 Then
        rs5.AddNew
        rs5.Fields("season").Value = Trim(sesson.Caption)
        rs5.Fields("SL_NO").Value = Trim(Text1.Text)
        rs5.Fields("date_in").Value = CDate(DTPicker1.Value)
        rs5.Fields("time_in").Value = Trim(intime.Caption)
        rs5.Fields("TC_CODE").Value = Trim(cco.Caption)
        rs5.Fields("V_no").Value = Trim(vno.Caption)
        rs5.Fields("o_name").Value = Trim(op1.Caption)
        rs5.Fields("Tm_CODE").Value = Trim(chno.Caption)
        rs5.Fields("first_wt").Value = Val(fwt.Caption)
        rs5.Fields("RLW").Value = Val(rlw.Caption)
        rs5.Fields("do_no").Value = dono.Caption
        rs5.Fields("coll_code").Value = collcode.Caption
        rs5.Fields("order_qty").Value = Val(oqty.Caption)
        rs5.Fields("dest").Value = cadr.Caption
        rs5.Fields("shift_in").Value = Trim(inshift.Caption)
        rs5.Fields("tag").Value = ""
    End If
    rs5.Fields("o2_name").Value = op2.Caption
    rs5.Fields("time_out").Value = outtime.Caption
    rs5.Fields("date_out").Value = CDate(odate.Caption)
    rs5.Fields("second_wt").Value = Val(Text2.Text)
    rs5.Fields("balance_qty").Value = Val(bqty.Caption)
    rs5.Fields("challan_date").Value = DTPicker4.Value
    rs5.Fields("challan_no").Value = Trim(Text18.Text)
    rs5.Update
    rs5.Close
End If

    dataset
    rs1.Fields("o2_name").Value = op2.Caption
    rs1.Fields("time_out").Value = outtime.Caption
    rs1.Fields("date_out").Value = CDate(odate.Caption)
    rs1.Fields("second_wt").Value = Val(Text2.Text)
    rs1.Fields("balance_qty").Value = Val(bqty.Caption)
    rs1.Fields("challan_date").Value = DTPicker4.Value
    rs1.Fields("challan_no").Value = Trim(Text18.Text)
    rs1.Update
    
    invoicerep1

    If writefile = False Then
        MsgBox "File could not be transfered, try to save again"
        Command1.Enabled = True
        Command2.Enabled = True
        Exit Sub
    End If
        
    'If gpscode = 1 Then
'        writeora = False
'        writeserver
'        If writeora = False Then
'            MsgBox "File could not be transfered to Oracle server"
'        Else
'            Set rs1 = New ADODB.Recordset
'            rs1.Open "Select * from special where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
'            If rs1.RecordCount > 0 Then
'                rs1.MoveFirst
'                rs1.Fields("data38") = "1"
'                rs1.Update
'            End If
'            rs1.Close
'            MsgBox "Data transfered to oracle server successfully"
'        End If
    'End If
        
    b = 1
    While b = 1
    b = MsgBox("Print Slip ", vbOKCancel, "Print ?")
    If b = 1 Then
        'printdata
        'If camcode = 0 Then
        '    dosprint
        'Else
            printop = 1
            printrep
        'End If
    End If
    Wend
    
    'printinv
    b = 1
    While b = 1
    b = MsgBox("Print Invoice ?", vbOKCancel, "Print ?")
    If b = 1 Then
        'dosprint
        printop = 1
        invoicerep
    End If
    Wend
    
    tctr = 0
If rfidcode = 1 And boomcode = 1 Then
If doorin = 1 Then
        openclose1
        d2 = 1
        door2.FillColor = &HFF00&
Else
        openclose
        d1 = 1
        door1.FillColor = &HFF00&
End If
End If
Else
MsgBox "Transaction Cancelled..2nd weighment again"
    
End If
Command1.Enabled = False
Command2.Enabled = True
Exit Sub

errm:
MsgBox Err.Description + "...tag not updated"
Command1.Enabled = True
Command2.Enabled = True
Exit Sub

errm1:
MsgBox Err.Description + "...Second weight not saved"
Command2.Enabled = True
Command1.Enabled = True
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
excdata(6) = rs2.Fields("destination")
wstream.WriteLine PADR("Sel.Exc.Reg.No.: ", 16) & PADR(rs4.Fields("u_excno"), 23) & PADR(" Pur.Exc.RegNo.: ", 17) & PADR(rs2.Fields("exc_reg_no"), 24)
wstream.WriteLine PADR("Range: ", 16) & PADR(rs4.Fields("u_range"), 23) & PADR(" Range: ", 17) & PADR(rs2.Fields("range"), 24)
wstream.WriteLine PADR("Division: ", 16) & PADR(rs4.Fields("u_division"), 23) & PADR(" Division: ", 17) & PADR(rs2.Fields("division"), 24)
wstream.WriteLine PADR("Commissionerate: ", 16) & PADR(rs4.Fields("u_commissionrate"), 23) & PADR(" Commissionerate: ", 17) & PADR(rs2.Fields("commissionerate"), 24)
wstream.WriteLine PADR("Vat Tin No.: ", 16) & PADR(rs4.Fields("u_tinno"), 23) & PADR(" Vat Tin No.: ", 17) & PADR(rs2.Fields("vat_tin_no"), 24)
wstream.WriteLine PADR("CST No.: ", 16) & PADR(rs4.Fields("u_cstno"), 23) & PADR(" CST No.: ", 17) & PADR(rs2.Fields("cst_no"), 24)
wstream.WriteLine PADR("PAN No.: ", 16) & PADR(rs4.Fields("u_panno"), 23) & PADR(" PAN No.: ", 17) & PADR(rs2.Fields("PAN"), 24)
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
wstream.WriteLine PADC(odate, 15) & PADC(Text18.Text, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
wstream.WriteLine String(80, "-")

wstream.WriteLine PADR("Tot.Lifted Qty:", 15) + PADR(LTrim(lqty.Caption), 10) & PADR("BILLED HEAD", 30) & padl("RATE", 12) & padl("VALUE", 13)
wstream.WriteLine String(25, " ") & PADR("", 30) & String(25, "-")
wstream.WriteLine String(25, " ") & PADR("Basic Value", 30) & padl(Format(CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 12) & padl(Format(lifted * CDbl(rs2.Fields("basic_rate")), "#####0.#0"), 13)
excdata(8) = rs2.Fields("basic_rate")
excdata(9) = lifted * CDbl(rs2.Fields("basic_rate"))
wstream.WriteLine String(25, " ") & PADR("Add on Price (WRC)", 30) & padl(CDbl(rs2.Fields("wrc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("wrc")), "#####0.#0"), 13)
excdata(10) = rs2.Fields("wrc")
excdata(11) = lifted * CDbl(rs2.Fields("wrc"))
wstream.WriteLine String(25, " ") & PADR("Selective Loading Charge", 30) & padl(CDbl(rs2.Fields("slc")), 12) & padl(Format(lifted * CDbl(rs2.Fields("slc")), "#####0.#0"), 13)
excdata(12) = rs2.Fields("slc")
excdata(13) = lifted * CDbl(rs2.Fields("slc"))
wstream.WriteLine PADR("Item Code: 27011910", 25) & PADR("Weighment Charge", 30) & padl(CDbl(rs2.Fields("weighment_chg")), 12) & padl(Format(lifted * CDbl(rs2.Fields("weighment_chg")), "#####0.#0"), 13)
excdata(14) = rs2.Fields("weighment_chg")
excdata(15) = lifted * CDbl(rs2.Fields("weighment_chg"))
excdata(16) = "0.00"
excdata(17) = "0.00"

excgross = (lifted * CDbl(rs2.Fields("basic_rate"))) + (lifted * CDbl(rs2.Fields("wrc"))) + (lifted * CDbl(rs2.Fields("slc"))) + (lifted * CDbl(rs2.Fields("weighment_chg"))) + (lifted * CDbl(rs2.Fields("royalty"))) + (lifted * CDbl(rs2.Fields("sed")))
excdata(18) = excgross
wstream.WriteLine PADR("Description of Goods", 30) & PADR("Excisable Gross", 32) & padl(Format(excgross, "#####0.#0"), 13)
wstream.WriteLine PADR("Raw Coal", 30) & PADR("Cent Excise @" + rs2.Fields("cent_exc_rate") + "%", 32) & padl(Format(excgross * Val(rs2.Fields("cent_exc_rate")) * 0.01, "#####0.#0"), 13)
wstream.WriteLine PADR("Destination: " & Trim(cadr), 30) & PADR("Ed Cess on Excise @" + rs2.Fields("edu_cess_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("edu_cess_rate")) * 0.01, "#####0.#0"), 13)
wstream.WriteLine PADR(" ", 30) & PADR("S/H Cess on Excise @" + rs2.Fields("high_edu_rate") + "%", 32) & padl(Format(excgross * 0.06 * Val(rs2.Fields("high_edu_rate")) * 0.01, "#####0.#0"), 13)
wstream.WriteLine PADR("Type of Coal:", 25) & PADR("Royalty", 30) & PADR(CDbl(rs2.Fields("royalty")), 12) & padl(Format(lifted * CDbl(rs2.Fields("royalty")), "#####0.#0"), 13)
excdata(26) = rs2.Fields("royalty")
excdata(27) = lifted * CDbl(rs2.Fields("royalty"))
wstream.WriteLine PADR("Non-Coking Coal", 25) & PADR("Stowing Excise Duty(SED)", 30) & padl(CDbl(rs2.Fields("sed")), 12) & padl(Format(lifted * CDbl(rs2.Fields("sed")), "#####0.#0"), 13)
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

wstream.WriteLine PADR("", 25) & PADR("Total Excise Duty", 42) & padl(Format(excduty, "#####0.#0"), 13)
decpart = Right(Format(excduty, "######.#0"), 2)
excdut1 = Left(Format(excduty, "0#####.#0"), Len(Format(excduty, "0#####.#0")) - 3)
wstream.WriteLine "Rupees " & worda(Val(excdut1)) & " and " & worda(CLng(decpart)) & " paise only"
wstream.WriteLine String(80, "-")
wstream.WriteLine String(25, " ") & PADR("Clean Energy Cess", 30) & padl(CDbl(rs2.Fields("clean_engy_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("clean_engy_cess")), "#####0.#0"), 13)
excdata(30) = rs2.Fields("clean_engy_cess")
excdata(31) = lifted * CDbl(rs2.Fields("clean_engy_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("clean_engy_cess"))
wstream.WriteLine String(25, " ") & PADR("Road/ RE Cess", 30) & padl(CDbl(rs2.Fields("road_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("road_cess")), "#####0.#0"), 13)
excdata(32) = rs2.Fields("road_cess")
excdata(33) = lifted * CDbl(rs2.Fields("road_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("road_cess"))
wstream.WriteLine String(25, " ") & PADR("PWD Cess/ MADA(Bazaar)", 30) & padl(CDbl(rs2.Fields("bazar_fee")), 12) & padl(Format(lifted * CDbl(rs2.Fields("bazar_fee")), "#####0.#0"), 13)
excdata(34) = rs2.Fields("bazar_fee")
excdata(35) = lifted * CDbl(rs2.Fields("bazar_fee"))
totgross = totgross + (lifted * CDbl(rs2.Fields("bazar_fee")))
wstream.WriteLine String(25, " ") & PADR("AMBH Cess", 30) & padl(CDbl(rs2.Fields("ambh_cess")), 12) & padl(Format(lifted * CDbl(rs2.Fields("ambh_cess")), "#####0.#0"), 13)
excdata(36) = rs2.Fields("ambh_cess")
excdata(37) = lifted * CDbl(rs2.Fields("ambh_cess"))
totgross = totgross + lifted * CDbl(rs2.Fields("ambh_cess"))
wstream.WriteLine String(25, " ") & PADR("Other Charges", 30) & padl(" ", 12) & padl(CDbl(rs2.Fields("other_charges")), 13)
excdata(38) = rs2.Fields("other_charges")
excdata(39) = lifted * CDbl(rs2.Fields("other_charges"))
totgross = totgross + (lifted * CDbl(rs2.Fields("other_charges")))
excdata(40) = totgross
wstream.WriteLine String(25, " ") & String(55, "-")
wstream.WriteLine String(25, " ") & PADR("TOTAL VALUE", 42) & padl(Format(totgross, "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & PADR("VAT/CST@ " + rs2.Fields("tax_percent") + "%", 30) & padl("JST", 12) & padl(Format(totgross * 0.01 * CDbl(rs2.Fields("tax_percent")), "#####0.#0"), 13)
excdata(41) = rs2.Fields("tax_percent")
excdata(42) = totgross * 0.01 * CDbl(rs2.Fields("tax_percent"))
tcs = totgross * 0.01
totgross = totgross + (totgross * 0.01 * CDbl(rs2.Fields("tax_percent")))
excdata(43) = totgross
wstream.WriteLine String(25, " ") & PADR("TCS@ 1%", 30) & padl(" ", 12) & padl(Format(tcs, "#####0.#0"), 13)
excdata(44) = "1"
excdata(45) = tcs
wstream.WriteLine String(25, " ") & String(55, "-")
totgross = totgross + tcs
excdata(46) = totgross
excdata(47) = "2/" & arcode & wbcode & "/" & Trim(cco.Caption) & "/" & dono.Caption & "/" & padl(Format(dsno, "0####0"), 6)
excdata(48) = odate.Caption
excdata(49) = rtype.Caption

wstream.WriteLine String(25, " ") & PADR("Gross Value", 42) & padl(Format(totgross, "#####0.#0"), 13)
wstream.WriteLine String(25, " ") & String(55, "-")
wstream.WriteLine "Gross value in words:"
decpart = Right(Format(totgross, "######.#0"), 2)
totgross = 0
On Error Resume Next
totgross = Left(Format(totgross, "######.#0"), Len(Format(totgross, "######.#0")) - 3)
wstream.WriteLine "Rupees " & worda(Val(totgross)) & "and " & worda(CLng(decpart)) & "paise only"
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
invrep.Sections("Section1").Controls("Label10").Caption = "TAG No : " & Trim(Text3.Text)

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

'dd1amt = rs2.Fields("draft_amt1")
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
invrep.Sections("Section1").Controls("Label35").Caption = PADC(odate, 15) & PADC(Text18.Text, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
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
totgross = 0
invrep.Sections("Section1").Controls("Label66").Caption = String(80, "-")
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
invrep.Sections("Section1").Controls("Label10").Caption = "TAG No : " & Trim(Text3.Text)

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

'dd1amt = rs2.Fields("draft_amt1")
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
invrep.Sections("Section1").Controls("Label35").Caption = PADC(odate, 15) & PADC(Text18.Text, 15) & PADC(vno.Caption, 15) & PADC(rs2.Fields("grade"), 15) & padl(Format(lifted, "#####0.#0"), 15)
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
totgross = 0
invrep.Sections("Section1").Controls("Label66").Caption = String(80, "-")
invrep.Sections("Section1").Controls("Label67").Caption = PADR("Prepared By", 26) & PADR("Checked By", 26) & PADR("Auth. Signatory", 26)
invrep.Sections("Section1").Controls("Label68").Caption = PADR(" ", 45) & PADR("Under Jurisdiction of Jharkhand", 35)

rs2.Close
rs4.Close
End Sub


Private Sub printrep()
Dim ordst As String
Dim orden As String
On Error Resume Next
spcsndwt.Sections(1).Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
spcsndwt.Sections(1).Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
spcsndwt.Sections(1).Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)

spcsndwt.Sections(1).Controls("Label4").Caption = PADC("Weighment Slip", 25)

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

spcsndwt.Sections(1).Controls("Label9").Caption = "TAG No : " & Left(Text3.Text, 4) & "XXXXXXXXXX" & Right(Text3.Text, 4)
spcsndwt.Sections(1).Controls("Label10").Caption = PADR("Operator 1 ", 20) & padl(": ", 2) & rs1.Fields("O_name").Value & "      Operator 2 " & padl(": ", 2) & rs1.Fields("O2_name").Value

spcsndwt.Sections(1).Controls("Label11").Caption = PADR("Vehicle No", 20) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 15)

spcsndwt.Sections(1).Controls("Label12").Caption = PADR("Name", 20) & Chr(27) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
spcsndwt.Sections(1).Controls("Label13").Caption = PADR("Destination", 20) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)
spcsndwt.Sections(1).Controls("Label14").Caption = PADR("Order No.", 20) & padl(":", 2) & padl(rs1.Fields("DO_NO").Value, 35)

ordst = Mid(rs1.Fields("do_start_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 1, 4)
orden = Mid(rs1.Fields("do_end_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 1, 4)
spcsndwt.Sections(1).Controls("Label15").Caption = PADR("Order Validity", 20) & padl(":", 2) & padl(ordst, 10) & " - " & padl(orden, 10)

spcsndwt.Sections(1).Controls("Label16").Caption = PADR("Colliery Name", 20) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
spcsndwt.Sections(1).Controls("Label17").Caption = PADR("Material", 20) & padl(":", 2) & padl(chno.Caption, 35)
spcsndwt.Sections(1).Controls("Label18").Caption = PADR("Hologram No.", 20) & padl(":", 2) & padl(Text18.Text, 35)

spcsndwt.Sections(1).Controls("Label19").Caption = "Order Qty: " & oqty & "   Lifted Qty: " & lqty & "   Balance Qty: " & bqty
spcsndwt.Sections(1).Controls("Label20").Caption = padl("Date In", 20) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & padl("Date Out", 10) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
spcsndwt.Sections(1).Controls("Label21").Caption = padl("Time In", 20) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & padl("Time Out", 10) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)

spcsndwt.Sections(1).Controls("Label22").Caption = PADR("RLW", 20) & padl(":", 2) & padl(rs1.Fields("RLW").Value, 15) & "Kg"

spcsndwt.Sections(1).Controls("Label23").Caption = PADR("First Weight", 14) & padl(str(rs1.Fields("First_Wt").Value) + " Kg", 35) & Chr(27)
spcsndwt.Sections(1).Controls("Label24").Caption = PADR("Second Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15)
spcsndwt.Sections(1).Controls("Label25").Caption = PADR("Net Weight   ", 13) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15)


Else
spcsndwt.Sections(1).Controls("Label5").Caption = PADC(Trim("No Such Serial No Created"), 30)
End If
If printop = 1 Then
spcsndwt.Show vbModal
End If
End Sub


Private Sub printdata()
Dim ordst As String
Dim orden As String
Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("Weighment Slip", 28) & Chr(27) & "F"
 wstream.WriteLine Chr(18) & String(53, "-")
Set rs1 = New ADODB.Recordset
rs1.Open "sELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
'wstream.WriteLine Chr(18) & "E" & Chr(14) & padl("Year   ", 10) & Chr(27) & "F" & Chr(18) & "E" & padl(rs1.Fields("SEASON").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("SL_NO").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("1st Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("2nd Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O2_name").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name    ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Destination ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("do_no").Value, 35)
ordst = Mid(rs1.Fields("do_start_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_start_date").Value, 1, 4)
orden = Mid(rs1.Fields("do_end_date").Value, 7, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 5, 2) + "/" + Mid(rs1.Fields("do_end_date").Value, 1, 4)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order Validity", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(ordst, 15) & " - " & padl(orden, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colliery Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Material", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(chno.Caption, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Challan No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Challan_No").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & "Order Qty: " & oqty & "     Lifted Qty: " & lqty & "     Balance Qty: " & bqty
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Date Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_out").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15) & Chr(18) & Chr(27) & "E" & padl("Time Out", 10) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_out").Value, 15)

wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("RLW", 13) & padl(":", 2) & padl(str(rs1.Fields("RLW").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("First_Wt").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Second Weight", 13) & padl(":", 2) & padl(str(rs1.Fields("second_Wt").Value) & " Kg", 15) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("Net Weight   ", 13) & padl(":", 2) & padl(str(Abs(rs1.Fields("First_Wt").Value - rs1.Fields("SECOND_WT").Value)) & " Kg", 15) & Chr(27) & "F"
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
Command1.Enabled = False
Command2.Enabled = False
If rfidcode = 1 And boomcode = 1 Then
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
Else
Unload Me
End If
End Sub

Private Sub Command3_Click()

On Error Resume Next
Dim ax, bx As Boolean
cctv.DHSurveillanceCtrl1.EnablePreview (True)

Picture1.PaintPicture SaveFormPic, 0, 0, , , 100, 650, 3700, 2900
Picture1.PaintPicture SaveFormPic, 0, 0, , , 100, 650, 3700, 2900
Picture1.PaintPicture SaveFormPic, 0, 0, , , 100, 650, 3700, 2900

Picture2.PaintPicture SaveFormPic, 0, 0, , , 3700, 650, 3700, 2900
Picture2.PaintPicture SaveFormPic, 0, 0, , , 3700, 650, 3700, 2900
Picture2.PaintPicture SaveFormPic, 0, 0, , , 3700, 650, 3700, 2900

Picture3.PaintPicture SaveFormPic, 0, 0, , , 100, 3500, 3700, 2900
Picture3.PaintPicture SaveFormPic, 0, 0, , , 100, 3500, 3700, 2900
Picture3.PaintPicture SaveFormPic, 0, 0, , , 100, 3500, 3700, 2900
Picture4.PaintPicture SaveFormPic, 0, 0, , , 3700, 3500, 3700, 2900
Picture4.PaintPicture SaveFormPic, 0, 0, , , 3700, 3500, 3700, 2900
Picture4.PaintPicture SaveFormPic, 0, 0, , , 3700, 3500, 3700, 2900

SavePicture Picture1.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw1.jpg"
SavePicture Picture2.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw2.jpg"
SavePicture Picture3.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw3.jpg"
SavePicture Picture4.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw4.jpg"
End Sub


Private Sub Command4_Click()
If Trim(Texttag.Text) <> "" Then
    datashow
    'Text18.SetFocus
Else
    MsgBox "Invalid tag"
End If
End Sub


'Private Sub Command3_Click()
'On Error Resume Next
'Picture1.PaintPicture SaveFormPic, 10, -970, 17000, 12900
'Picture2.PaintPicture SaveFormPic, -3700, -970, 17000, 12900
'End Sub



'Private Sub Command4_Click()
'SavePicture Picture1.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw1.jpg"
'SavePicture Picture2.Image, "d:\wbdata\imagedata\" & Text1.Text & "_sw2.jpg"
'MsgBox "Image Saved"
'End Sub


Private Sub Command6_Click()
If timecheck(cco.Caption) = False Then
    MsgBox "Weighment not allowed..Invalid Time Range"
    Command6.Enabled = False
    Command1.Enabled = False
End If

If camcode = 1 Then
    Command3_Click
    DoEvents
    Command3_Click
    DoEvents
    spswt.SetFocus
End If
Text2.Text = main.comwt.Caption
Command6.Enabled = False
End Sub

Private Sub DTPicker4_Click()
If Command6.Enabled = True Then Command6.SetFocus
End Sub

Private Sub Form_Activate()
If Val(main.comwt.Caption) > 20 Then
    MsgBox "Platform not empty or not ready"
    Command1.Enabled = False
End If

FTP1.HostName = ips(2)
FTP1.username = unames(2)
FTP1.PassWord = passws(2)
FTP1.ChangeRemoteDir (paths(2))
If FTP1.Connect Then
Else
    MsgBox "Headquarter not connected..check connection and reload form", vbCritical
    Command1.Enabled = False
End If

'If gpscode = 1 Then
'    FTP2.HostName = "172.20.0.99"
'    FTP2.username = "arsftp1"
'    FTP2.PassWord = "B((1@r$f+p"
'    'FTP1.ChangeRemoteDir ("root")
'    If FTP2.Connect Then
'    Else
'        MsgBox "GPS server not connected", vbCritical
'    End If
'End If

Exit Sub

Command1.Enabled = False
doorin = 0
d1 = 0
d2 = 0
door1.FillColor = &HFF&
door2.FillColor = &HFF&
Label30.Caption = ""
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture

Conn
username = loginname
unlocktxt
DTPicker1.Value = Date
DTPicker4.Value = Date

If camcode = 1 Then
    Picture1.Visible = True
    Picture2.Visible = True
    Picture3.Visible = True
    Picture4.Visible = True
Else
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
End If

If boomcode = 1 Then
    Winsock1.Close
    Winsock1.RemotePort = paths(5)
    Winsock1.RemoteHost = ips(5)
    Winsock1.Connect

    Winsock2.Close
    Winsock2.RemotePort = paths(7)
    Winsock2.RemoteHost = ips(7)
    Winsock2.Connect
End If

On Error Resume Next
If rfidcode = 1 Then
a = Net_Connect(unames(6), CLng(passws(6)), ips(6), CLng(paths(6)))
Conn1
Texttag.Text = ""
tagvalid = False
Else
    Text1.Enabled = True
    Text1.Locked = False
End If
End Sub

Private Sub unlocktxt()
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
Text2.Text = Val(0)

op2.Caption = username
cco.Caption = ""
rlw.Caption = ""

collcode.Caption = ""
collname.Caption = ""
inshift.Caption = ""
outshift.Caption = ""
dono.Caption = ""
rtype.Caption = ""
''DTPicker2.Value = ""
'DTPicker3.Value = ""
oqty.Caption = ""
lqty.Caption = ""
bqty.Caption = ""
balqty = bqty.Caption
lifqty = lqty.Caption
Text1.Text = ""
'Text1.SetFocus



End Sub

Private Sub datashow()
Set rs1 = New ADODB.Recordset

If rfidcode = 1 Then
    rs1.Open "Select * from spdtwtrep where tag='" & Trim(Texttag.Text) & "' and second_wt=0 order by date_in desc", co, adOpenKeyset, adLockOptimistic
Else
    rs1.Open "Select * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "' and second_wt=0 order by date_in desc", co, adOpenKeyset, adLockOptimistic
End If

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    tctr1 = 0
    If doorin = 1 Then
        openclose
        d1 = 1
        door1.FillColor = &HFF00&
    ElseIf doorin = 2 Then
        openclose1
        d2 = 1
        door2.FillColor = &HFF00&
    End If
    
    rs1.MoveFirst
    op2.Caption = username
    sesson.Caption = rs1.Fields("season").Value
    vno.Caption = rs1.Fields("V_no").Value
    Text1.Text = rs1.Fields("sl_no").Value
    cco.Caption = rs1.Fields("custcd").Value
    cname.Caption = rs1.Fields("purchaser").Value
    cadr.Caption = rs1.Fields("dest").Value
    op1.Caption = rs1.Fields("o_name").Value
    
    intime.Caption = rs1.Fields("time_in").Value
    chno.Caption = rs1.Fields("tm_code").Value
    'material.Caption = rs1.Fields("m_name").Value
    fwt.Caption = rs1.Fields("first_wt").Value
    rlw.Caption = rs1.Fields("RLW").Value
    
    DTPicker1.Value = rs1.Fields("DATE_IN").Value
    collcode.Caption = rs1.Fields("coll_code").Value
    collname.Caption = rs1.Fields("coll_desc").Value
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


    If timecheck(cco.Caption) = False Then
        MsgBox "Weighment not allowed..Invalid Time Range"
        unlocktxt
        Command1.Enabled = False
    End If

    If Val(rs1.Fields("second_wt").Value & "") > 0 Then
    MsgBox "Dont try again ", vbCritical, "Second Weight already taken "
    unlocktxt
    Exit Sub
    End If
Else
'Timer2.Enabled = True
unlocktxt
DTPicker3.Value = Date
Text1.Text = ""
MsgBox "Serial Not Found", vbInformation, "Wrong Serial No"
End If
If DTPicker3.Value < Date - 3 Then
    MsgBox "Second weight cannot be taken 72 hrs after DO Validity"
    Text1.Text = ""
'    Text1.SetFocus
End If
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub


Private Sub Form_Terminate()
If boomcode = 1 Then
Winsock1.Close
Winsock2.Close
End If
main.texttagg.Caption = ""
main.tagg.Caption = ""
main.Timer3.Enabled = True
    validtag = False
main.unloadfrm
main.invalidtags.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
If boomcode = 1 Then
    Winsock1.Close
    Winsock2.Close
End If
main.texttagg.Caption = ""
main.tagg.Caption = ""
main.Timer3.Enabled = True
    validtag = False
main.unloadfrm
main.invalidtags.Clear
End Sub

Private Sub Text1_GotFocus()
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = forceno(KeyAscii)
If KeyAscii = 13 Then
    If Trim(Text1.Text) <> "" Then
        datashow
        'Text18.SetFocus
    End If
End If


End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'DTPicker4.SetFocus
'End If
End Sub

Private Sub Text2_Change()
ntw.Caption = Abs(Val(fwt.Caption) - Val(Text2.Text))
bqty.Caption = Val(balqty) - ntw
lqty.Caption = Val(lifqty) + ntw
If CLng(ntw) < 100 And Val(Text2.Text) > 0 Then
    ntw.Caption = 0
    Text2.Text = fwt.Caption
    MsgBox "Net weight too low, 0 weight taken", vbInformation, "Zero weight taken"
End If
'main.comwt.Caption = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text2.Text) = 0 Then
If Command6.Enabled = True Then Command6.SetFocus
Else
Command1.SetFocus
End If
End If

End Sub

Private Sub showtagdata()
On Error GoTo errm
Timer2.Enabled = False
Conn1
    Set rs4 = New ADODB.Recordset
    rs4.Open "Select * from tags where tagno = '" + Trim(Texttag) + "' and valid='2'", co1, adOpenKeyset, adLockOptimistic
    If rs4.RecordCount > 0 Then
        rs4.MoveFirst
        wtmode = rs4.Fields("mode").Value
        tagtrips = Val(rs4.Fields("tagtrips"))
        tripsdone = Val(rs4.Fields("trips_done")) + 1
        If Not IsNumeric(tripsdone) Then
        tripsdone = 1
        End If
    End If
    rs4.Close
    datashow
    
Exit Sub
errm:
    MsgBox Err.Description

End Sub







Private Sub Texttag_Change()
Text3.Text = Left(Texttag.Text, 4) + "XXXXXXXXXXXXXX" + Right(Texttag.Text, 4)
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

If wtmode = "D" And rfidcode = 1 Then
If main.comwt.Caption > 14000 Then
If doorin = 1 Then
    If d1 = 1 Then
        Timer5.Enabled = True
        If tctr1 = 15 Then
            openclose
            d1 = 0
            door1.FillColor = &HFF&
            tctr1 = 0
            Me.Timer5.Enabled = False
        End If
    End If
Else
    If d2 = 1 Then
        Timer5.Enabled = True
        If tctr1 = 15 Then
            openclose1
            d2 = 0
            door2.FillColor = &HFF&
            tctr1 = 0
            Me.Timer5.Enabled = False
        End If
    End If
End If
End If
End If

If wtmode = "R" And rfidcode = 1 Then
If main.comwt.Caption > 5000 Then
If doorin = 1 Then
    If d1 = 1 Then
        Timer5.Enabled = True
        If tctr1 = 15 Then
            openclose
            d1 = 0
            door1.FillColor = &HFF&
            tctr1 = 0
            Me.Timer5.Enabled = False
        End If
    End If
Else
    If d2 = 1 Then
        Timer5.Enabled = True
        If tctr1 = 15 Then
            openclose1
            d2 = 0
            door2.FillColor = &HFF&
            tctr1 = 0
            Me.Timer5.Enabled = False
        End If
    End If
End If
End If
End If


If main.comwt.Caption < 500 And rfidcode = 1 Then
If doorin = 1 Then
    If d2 = 1 Then
        Me.Timer3.Enabled = True
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
        Me.Timer3.Enabled = True
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

outtime.Caption = Format(Time, "hh:mm")
End Sub


Private Sub Timer2_Timer()
On Error Resume Next
If rfidcode = 1 Then
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
End If
End Sub

Private Sub Timer3_Timer()
If rfidcode = 1 And boomcode = 1 Then
tctr = tctr + 1
Label30.Caption = "GATE-OUT CLOSING: " + CStr(15 - tctr)
End If
End Sub

Private Sub Timer4_Timer()
If rfidcode = 1 And boomcode = 1 Then
tctr = tctr + 1
If tctr > 3 Then
Unload Me
End If
End If
End Sub

Private Sub Timer5_Timer()
If rfidcode = 1 And boomcode = 1 Then
tctr1 = tctr1 + 1
Label30.Caption = "GATE-IN CLOSING: " + CStr(15 - tctr1)
End If
End Sub
