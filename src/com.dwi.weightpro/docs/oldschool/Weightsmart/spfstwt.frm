VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form spfstwt 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3360
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   7680
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   7680
   End
   Begin VB.ComboBox readtags 
      Height          =   315
      Left            =   1920
      TabIndex        =   67
      Text            =   "Combo1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   7680
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   1440
      Top             =   7680
   End
   Begin MSComCtl2.DTPicker tmpdt 
      Height          =   375
      Left            =   720
      TabIndex        =   40
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Format          =   78381057
      CurrentDate     =   41165
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   7680
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2600
      Left            =   11400
      ScaleHeight     =   2535
      ScaleWidth      =   3555
      TabIndex        =   52
      Top             =   480
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7395
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   14895
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   4800
         TabIndex        =   71
         Top             =   5880
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   3840
         TabIndex        =   70
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   7680
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   56
         Top             =   360
         Width           =   3615
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   11280
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   54
         Top             =   2940
         Width           =   3615
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         Height          =   2600
         Left            =   7680
         ScaleHeight     =   2535
         ScaleWidth      =   3555
         TabIndex        =   53
         Top             =   2940
         Width           =   3615
      End
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1140
         Width           =   3255
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1140
         Width           =   375
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1140
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   9960
         TabIndex        =   55
         Top             =   5940
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
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
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Height          =   1635
         Left            =   10920
         TabIndex        =   47
         Top             =   5700
         Width           =   3975
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
            Height          =   1275
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmddel 
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
            Height          =   1275
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   240
            Width           =   2145
         End
         Begin VB.CommandButton cmdref 
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
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Visible         =   0   'False
            Width           =   1305
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   4200
         TabIndex        =   41
         Top             =   1740
         Width           =   3300
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordered Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   45
            Top             =   120
            Width           =   1365
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   1800
            TabIndex        =   44
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance Qty"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   105
            TabIndex        =   43
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   105
            TabIndex        =   42
            Top             =   840
            Width           =   3030
         End
      End
      Begin VB.TextBox Text17 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   5880
         Width           =   3015
      End
      Begin VB.TextBox Text16 
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
         Left            =   1800
         TabIndex        =   13
         Top             =   6240
         Width           =   3015
      End
      Begin VB.TextBox Text8 
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
         Left            =   7920
         TabIndex        =   11
         Text            =   "-na-"
         Top             =   3360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox Text7 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   5340
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   9
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Text5 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   3720
         Width           =   5535
      End
      Begin VB.TextBox Text10 
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
         Left            =   1800
         TabIndex        =   7
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text15 
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
         Left            =   3840
         TabIndex        =   6
         Top             =   2580
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text11 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2460
         Width           =   1170
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
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   6840
         Width           =   1575
      End
      Begin VB.TextBox Text9 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   6840
         Width           =   1455
      End
      Begin VB.TextBox Text4 
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2100
         Width           =   1695
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1740
         Width           =   2415
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   78381057
         CurrentDate     =   39961
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   1800
         TabIndex        =   38
         Top             =   4800
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
         Format          =   78381057
         CurrentDate     =   40669
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1800
         TabIndex        =   39
         Top             =   4440
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
         Format          =   78381057
         CurrentDate     =   40669
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3000
         TabIndex        =   69
         Top             =   2580
         Width           =   285
      End
      Begin VB.Label doorin1 
         Caption         =   "0"
         Height          =   195
         Left            =   5760
         TabIndex        =   66
         Top             =   6780
         Width           =   975
      End
      Begin VB.Label d21 
         Caption         =   "0"
         Height          =   315
         Left            =   6240
         TabIndex        =   65
         Top             =   7020
         Width           =   375
      End
      Begin VB.Label d11 
         Caption         =   "0"
         Height          =   315
         Left            =   5760
         TabIndex        =   64
         Top             =   7020
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
         Height          =   555
         Left            =   5280
         TabIndex        =   62
         Top             =   6000
         Width           =   4935
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
         Left            =   8640
         TabIndex        =   61
         Top             =   6660
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
         Left            =   7680
         TabIndex        =   60
         Top             =   6660
         Width           =   375
      End
      Begin VB.Shape door2 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   8640
         Top             =   6960
         Width           =   375
      End
      Begin VB.Shape door1 
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7680
         Top             =   6960
         Width           =   375
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   58
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Colliery Code"
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
         Left            =   240
         TabIndex        =   37
         Top             =   6000
         Width           =   1425
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Colliery"
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
         Left            =   240
         TabIndex        =   36
         Top             =   6360
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Record Type"
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
         Left            =   3840
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
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
         Left            =   240
         TabIndex        =   34
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label24 
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
         Left            =   240
         TabIndex        =   33
         Top             =   2580
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
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
         Left            =   240
         TabIndex        =   32
         Top             =   4560
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
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
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   1170
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
         Left            =   240
         TabIndex        =   29
         Top             =   6960
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Material"
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
         Left            =   8160
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   5460
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
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
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   1185
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
         Left            =   240
         TabIndex        =   25
         Top             =   3840
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
         Left            =   240
         TabIndex        =   24
         Top             =   2220
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Op Name"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1860
         Width           =   990
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
         Left            =   4080
         TabIndex        =   22
         Top             =   720
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Daily Sl No"
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
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1185
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
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP FTP2 
      Left            =   6240
      OleObjectBlob   =   "spfstwt.frx":0000
      Top             =   7680
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP FTP1 
      Left            =   5640
      OleObjectBlob   =   "spfstwt.frx":0024
      Top             =   7680
   End
End
Attribute VB_Name = "spfstwt"
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
Dim chkpass As Integer
Dim stadate As Date
Dim enddate As Date
Dim dallot As Integer
Private WithEvents cFTP As clsFTP
Attribute cFTP.VB_VarHelpID = -1
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
Dim tctr1 As Integer

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

Function writefile() As Boolean
Dim fname As String
Dim weightord As Double
Dim weightbal As Double
Dim weightdiff As Double
Dim ftime As String
Dim writestr As String

ftime = Left(Text2.Text, 2) + Right(Text2.Text, 2) + "00"
weightord = CDbl(Label20) / 1000
weightbal = CDbl(Label21) / 1000
weightdiff = (CDbl(Label20) - CDbl(Label21)) / 1000
fname = "tr" + Trim(Text11.Text) + Trim(Text1.Text) + ".txt"
writestr = padl(Trim(Text15.Text), 1) & "|" & padl(Trim(Text11.Text), 11) & "|" & trimdate(DTPicker3.Value) & "|" & trimdate(DTPicker4.Value) & "|" & padl(Trim(Text10.Text), 6) & "|" & padl(Trim(Text17.Text), 4) & "|" & padl(Trim(Text7), 10) & "|" & padl(Format(weightord, "0####.##0"), 9) & "|" & padl(Format(Trim(CStr(weightdiff)), "0####.##0"), 9) & "|" & padl(Format(weightbal, "0####.##0"), 9) & "|" & padl(Trim(Text1), 12) & "|" & padl(Format(CDbl(Text13) / 1000, "0####.##0"), 9) & "|" & padl(Trim(Text4), 15) & "|" & padl("", 8) & "|" & padl("", 8) & "|" & padl(trimdate(DTPicker1.Value), 8) & "|" & padl(Trim(ftime), 6) & "|" & padl("", 8) & "|" & padl("", 6) & "|" & padl("", 10) & "|" & padl(Format(CDbl(Text9) / 1000, "0#####.##0"), 10) & "|" & padl("", 10) & "|" & padl(Trim(Text6), 2) & "|" & padl(Trim(main.Label6), 1) & "|" & padl(Trim(Text3), 6) & "|" + vbCrLf

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


Function writegps() As Boolean
Dim fname As String
Dim weightord As Double
Dim weightbal As Double
Dim weightdiff As Double
Dim ftime As String
Dim writestr As String

ftime = Left(Text2.Text, 2) + Right(Text2.Text, 2) + "00"
weightord = CDbl(Label20) / 1000
weightbal = CDbl(Label21) / 1000
weightdiff = (CDbl(Label20) - CDbl(Label21)) / 1000
fname = "tr" + Trim(Text11.Text) + Trim(Text1.Text) + ".txt"
writestr = padl(Trim(Text15.Text), 1) & "|" & padl(Trim(Text11.Text), 11) & "|" & trimdate(DTPicker3.Value) & "|" & trimdate(DTPicker4.Value) & "|" & padl(Trim(Text10.Text), 6) & "|" & padl(Trim(Text17.Text), 4) & "|" & padl(Trim(Text7), 10) & "|" & padl(Format(weightord, "0####.##0"), 9) & "|" & padl(Format(Trim(CStr(weightdiff)), "0####.##0"), 9) & "|" & padl(Format(weightbal, "0####.##0"), 9) & "|" & padl(Trim(Text1), 12) & "|" & padl(Format(CDbl(Text13) / 1000, "0####.##0"), 9) & "|" & padl(Trim(Text4), 15) & "|" & padl("", 8) & "|" & padl("", 8) & "|" & padl(trimdate(DTPicker1.Value), 8) & "|" & padl(Trim(ftime), 6) & "|" & padl("", 8) & "|" & padl("", 6) & "|" & padl("", 10) & "|" & padl(Format(CDbl(Text9) / 1000, "0#####.##0"), 10) & "|" & padl("", 10) & "|" & padl(Trim(Text6), 2) & "|" & padl(Trim(main.Label6), 1) & "|" & padl(Trim(Text3), 6) & "|" + vbCrLf

writegps = True

If FTP2.Connect Then
    'If FTP2.ChangeRemoteDir(paths(2)) Then
        If FTP2.PutFileFromTextData(fname, writestr) Then
            writegps = True
            If FTP2.GetRemoteFileTextData(fname) = writestr Then
                MsgBox "File transfered fo gps server, Verification successful"
                writegps = True
            Else
                MsgBox "File verification failed..check gps connection"
                writegps = False
            End If
        Else
            MsgBox "cound not write file..check gps connection"
            writegps = False
        End If
    'Else
    '    MsgBox "gps directory not found..check connection and save again"
    '    writegps = False
    'End If
Else
    MsgBox "No connection to gps server"
    writegps = False
End If
End Function



Sub writefile_old()
Dim fname As String
Dim weightord As Double
Dim weightbal As Double
Dim weightdiff As Double
Dim fso As New FileSystemObject
Dim bSuccess As Boolean
Dim sError As String
Dim ftime As String
Dim Filist As String

ftime = Left(Text2.Text, 2) + Right(Text2.Text, 2) + "00"
weightord = CDbl(Label20) / 1000
weightbal = CDbl(Label21) / 1000
weightdiff = (CDbl(Label20) - CDbl(Label21)) / 1000
fname = "tr" + Trim(Text11.Text) + Trim(Text1.Text) + ".txt"
Set fsys = CreateObject("scripting.filesystemobject")

Set wstream = fsys.OpenTextFile("d:\wbdata\" & fname, ForWriting, True)
wstream.WriteLine padl(Trim(Text15.Text), 1) & "|" & padl(Trim(Text11.Text), 11) & "|" & trimdate(DTPicker3.Value) & "|" & trimdate(DTPicker4.Value) & "|" & padl(Trim(Text10.Text), 6) & "|" & padl(Trim(Text17.Text), 4) & "|" & padl(Trim(Text7), 10) & "|" & padl(Format(weightord, "0####.##0"), 9) & "|" & padl(Format(Trim(CStr(weightdiff)), "0####.##0"), 9) & "|" & padl(Format(weightbal, "0####.##0"), 9) & "|" & padl(Trim(Text1), 12) & "|" & padl(Format(CDbl(Text13) / 1000, "0####.##0"), 9) & "|" & padl(Trim(Text4), 15) & "|" & padl("", 8) & "|" & padl("", 8) & "|" & padl(trimdate(DTPicker1.Value), 8) & "|" & padl(Trim(ftime), 6) & "|" & padl("", 8) & "|" & padl("", 6) & "|" & padl("", 10) & "|" & padl(Format(CDbl(Text9) / 1000, "0#####.##0"), 10) & "|" & padl("", 10) & "|" & padl(Trim(Text6), 2) & "|" & padl(Trim(main.Label6), 1) & "|" & padl(Trim(Text3), 6) & "|" + vbCrLf
wstream.Close
SetAttr "d:\wbdata\" & fname, vbReadOnly
fso.CopyFile "d:\wbdata\" & fname, "d:\wbdata\areadata\" & fname

End Sub



Private Sub dataset()
recfound = False
Set rs1 = New ADODB.Recordset
 rs1.Open "Select * from  special where date_in = #" & Format(CDate(DTPicker1.Value), "mm/dd/yyyy") & "# and sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
recfound = True
End If
End Sub

Private Sub cmddel_Click()
cmdsave.Enabled = False
cmddel.Enabled = False

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

Private Sub cmdref_Click()
unlocktxt
Label20 = "0"
Label21 = "0"
Label23 = ""
Text4.SetFocus
End Sub

Private Sub cmdsave_Click()
Dim costr As String
Dim costr1 As String

cmddel.Enabled = False
cmdsave.Enabled = False
If Val(Label21.Caption) >= Abs(chkwt1 - chkwt2) Then
If Val(Text9.Text) < chkwt1 Then
    MsgBox "Weight is less than minimum weight", vbInformation, "Weight is too low "
    Exit Sub
End If

If Val(Text9.Text) > (Text13.Text) Then
    MsgBox "Weight is more than RLW...Not saved", vbInformation, "Weight is too high "
    cmdsave.Enabled = False
    Exit Sub
End If

a = MsgBox("Want to Update Records", vbOKCancel, "Update ?")
If a = 1 Then
'Command4_Click
dataset

If writefile = False Then
    MsgBox "File could not be transfered, try to save again"
    cmdsave.Enabled = True
    cmddel.Enabled = True
    Exit Sub
End If

    If gpscode = 1 Then
    If writegps = False Then
        MsgBox "File could not be transfered to GPS server"
    End If
    End If
    
On Error Resume Next
Conn2
If rfidcode = 1 Then
    Conn1
    costr = "insert into special(season,sl_no,date_in,time_in,tc_code,v_no,o_name,tm_code,first_wt,second_wt,rlw,do_no,coll_code,order_qty,balance_qty,dest,shift_in,tag)" _
    + " values ('" + Trim(Combo1.Text) + "','" + Trim(Text1.Text) + "','" + Format(Year(DTPicker1.Value), "00##") _
    + Format(Month(DTPicker1.Value), "0#") + Format(Day(DTPicker1.Value), "0#") + "','" + Trim(Text2.Text) + "','" _
    + Trim(Text10.Text) + "','" + Trim(Text4.Text) + "','" + Trim(Text3.Text) + "','" + Trim(Text7.Text) + "'," _
    + Trim(Text9.Text) + ",0," + Trim(Text13.Text) + ",'" + Trim(Text11.Text) + "','" + Trim(Text17.Text) + "'," _
    + Trim(Label20) + "," + Trim(Label21) + ",'" + Text6.Text + "','" + Trim(main.Label6.Caption) + "','" _
    + Trim(Text12.Text) + "')"
    co1.Execute costr
    
    On Error GoTo errm
    costr1 = "update tags set valid='2',tsno='" + Trim(Text1.Text) + "' where tagno = '" + Trim(Text12.Text) + "' and valid='1'"
    co1.Execute costr1

Else
    Set rs5 = New ADODB.Recordset
    rs5.Open "Select * from  specialtmp where date_in = #" & Format(CDate(DTPicker1.Value), "mm/dd/yyyy") & "# and sl_no='" & Trim(Text1.Text) & "'", co2, adOpenKeyset, adLockOptimistic
    If rs5.RecordCount = 0 Then
        rs5.AddNew
    End If
    
    rs5.Fields("season").Value = Trim(Combo1.Text)
    rs5.Fields("SL_NO").Value = Trim(Text1.Text)
    rs5.Fields("date_in").Value = CDate(DTPicker1.Value)
    rs5.Fields("time_in").Value = Trim(Text2.Text)
    rs5.Fields("TC_CODE").Value = Trim(Text10.Text)
    rs5.Fields("V_no").Value = Trim(Text4.Text)
    rs5.Fields("o_name").Value = Trim(Text3.Text)
    rs5.Fields("Tm_CODE").Value = Trim(Text7.Text)
    rs5.Fields("first_wt").Value = Val(Text9.Text)
    rs5.Fields("second_wt").Value = 0
    rs5.Fields("RLW").Value = Val(Text13.Text)
    rs5.Fields("do_no").Value = Text11.Text
    rs5.Fields("coll_code").Value = Text17.Text
    rs5.Fields("order_qty").Value = Val(Label20)
    rs5.Fields("balance_qty").Value = Val(Label21)
    rs5.Fields("dest").Value = Text6.Text
    rs5.Fields("shift_in").Value = Trim(main.Label6.Caption)
    rs5.Fields("tag").Value = Trim(Text12.Text)
    rs5.Update
    rs5.Close

End If

If recfound = False Then
rs1.AddNew
End If

rs1.Fields("season").Value = Trim(Combo1.Text)
rs1.Fields("SL_NO").Value = Trim(Text1.Text)
rs1.Fields("date_in").Value = CDate(DTPicker1.Value)
rs1.Fields("time_in").Value = Trim(Text2.Text)
rs1.Fields("TC_CODE").Value = Trim(Text10.Text)
rs1.Fields("V_no").Value = Trim(Text4.Text)
rs1.Fields("o_name").Value = Trim(Text3.Text)
rs1.Fields("Tm_CODE").Value = Trim(Text7.Text)
rs1.Fields("first_wt").Value = Val(Text9.Text)
rs1.Fields("second_wt").Value = 0
rs1.Fields("RLW").Value = Val(Text13.Text)
'MsgBox Text13.Text
rs1.Fields("do_no").Value = Text11.Text
rs1.Fields("coll_code").Value = Text17.Text
rs1.Fields("order_qty").Value = Val(Label20)
rs1.Fields("balance_qty").Value = Val(Label21)
rs1.Fields("dest").Value = Text6.Text
rs1.Fields("shift_in").Value = Trim(main.Label6.Caption)
rs1.Fields("tag").Value = Trim(Text12.Text)
rs1.Update
rs1.Close

recfound = False

b = 1
While b = 1
b = MsgBox("Print Slip", vbOKCancel, "Print ?")
If b = 1 Then
    'printdata
    'If camcode = 0 Then
    '    dosprint
    'Else
        printrep
    'End If
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

End If
Else
    MsgBox "Weighment not possible, order balance quantity less than " & CStr(Abs(chkwt1 - chkwt2)), , "Check weight"
End If
cmdsave.Enabled = False
cmddel.Enabled = True
Exit Sub

errm:
MsgBox Err.Description + "..Weighment not saved"
cmdsave.Enabled = True
cmddel.Enabled = True
End Sub


Private Sub printrep()
On Error Resume Next
spcfstwt.Sections(1).Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
spcfstwt.Sections(1).Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
spcfstwt.Sections(1).Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)

spcfstwt.Sections(1).Controls("Label4").Caption = PADC("Weighment Slip cum Loading Advice", 35)

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set spcfstwt.DataSource = rs1
spcfstwt.Sections(1).Controls("Text1").DataField = "SEASON"
Set spcfstwt.Sections(1).Controls("Image1").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_fw1.jpg")
Set spcfstwt.Sections(1).Controls("Image2").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_fw2.jpg")
Set spcfstwt.Sections(1).Controls("Image3").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_fw3.jpg")
Set spcfstwt.Sections(1).Controls("Image4").Picture = LoadPicture("d:\wbdata\imagedata\" & Text1.Text & "_fw4.jpg")
spcfstwt.Sections(1).Controls("Label5").Caption = padl("Year   ", 20) & padl(":", 2) & padl(rs1.Fields("SEASON").Value, 15)
spcfstwt.Sections(1).Controls("Label6").Caption = padl("Date   ", 20) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15)
spcfstwt.Sections(1).Controls("Label7").Caption = padl("Time In", 20) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15)
spcfstwt.Sections(1).Controls("Label8").Caption = PADR("Daily Serial", 20) & padl(":", 2) & padl(rs1.Fields("sl_no").Value, 15)

spcfstwt.Sections(1).Controls("Label9").Caption = PADR("Operator 1 ", 12) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 15) & "   TAG: " & Trim(Text18.Text)

spcfstwt.Sections(1).Controls("Label10").Caption = PADR("Vehicle No", 20) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 15)

spcfstwt.Sections(1).Controls("Label11").Caption = PADR("Name", 20) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
spcfstwt.Sections(1).Controls("Label12").Caption = PADR("Destination", 20) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)

spcfstwt.Sections(1).Controls("Label13").Caption = PADR("Order No.", 20) & padl(":", 2) & padl(rs1.Fields("DO_NO").Value, 35)
spcfstwt.Sections(1).Controls("Label14").Caption = PADR("Order Validity", 20) & padl(":", 2) & Format(DTPicker3.Value, "dd/mm/yyyy") & " - " & Format(DTPicker4.Value, "dd/mm/yyyy")
spcfstwt.Sections(1).Controls("Label21").Caption = PADR("Order Qty : ", 20) & Trim(Label20.Caption) & PADR("Balance Qty : ", 15) & Trim(Label21.Caption)
spcfstwt.Sections(1).Controls("Label15").Caption = PADR("Colliery Name", 20) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
spcfstwt.Sections(1).Controls("Label16").Caption = PADR("Material", 20) & padl(":", 2) & padl(Text7.Text, 35)
spcfstwt.Sections(1).Controls("Label17").Caption = PADR("RLW", 20) & padl(":", 2) & padl(rs1.Fields("RLW").Value, 15) & "Kg"
spcfstwt.Sections(1).Controls("Label18").Caption = PADR("First Weight", 20) & padl(":", 2) & padl(str(rs1.Fields("First_Wt").Value) + " Kg", 35) & Chr(27) & "F"
'spcfstwt.Sections(1).Controls("Label19").Caption = PADR(COMPAUTH, 20) & padl(" ", 30) & padl("Weighing Operator", 30)
'spcfstwt.Sections(1).Controls("Label20").Caption = padl(compdes, 100)
Else
spcfstwt.Sections(1).Controls("Label5").Caption = PADC(Trim("No Such Serial No Created"), 30)
End If

spcfstwt.Show vbModal
End Sub


Private Sub printdata()
Set wstream = fsys.OpenTextFile(App.Path + "\" & "rep.txt", ForWriting, True)
wstream.WriteLine Chr(18) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) + PADC(Trim(main.Label1.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label2.Caption), 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim(main.Label3.Caption), 25) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(27) & "E" & Chr(14) & PADC("Weighment Slip", 25) & Chr(27) & "F"
wstream.WriteLine Chr(18) & String(53, "-")

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from spdtwtrep where sl_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst

wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Year   ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("SEASON").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Date   ", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("date_in").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & padl("Time In", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Time_in").Value, 15)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Daily Serial", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("sl_no").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Operator Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("O_name").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Vehicle No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("v_no").Value, 15)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("purchaser").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Destination", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("dest").Value, 35)

wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order No.", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("DO_NO").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Order Validity", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & Format(DTPicker3.Value, "dd/mm/yyyy") & " - " & Format(DTPicker4.Value, "dd/mm/yyyy")
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Colliery Name", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("coll_desc").Value, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Material", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(Text7.Text, 35)
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("RLW", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("RLW").Value, 15) & "Kg"
'wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR("Challan No", 20) & Chr(27) & "F" & Chr(18) & padl(":", 2) & padl(rs1.Fields("Challan_No").Value, 15)

wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & PADR("First Weight", 14) & padl(str(rs1.Fields("First_Wt").Value) + " Kg", 35) & Chr(27) & "F"
wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & String(48, "-")
'wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & Chr(27) & "E" & PADR(COMPAUTH, 20) & Chr(27) & "F" & padl(" ", 2) & padl("Weighing Operator", 30)
'wstream.WriteBlankLines 1
wstream.WriteLine Chr(18) & String(53, "-")

wstream.WriteLine Chr(15) & padl(compdes, 100)
Else
wstream.WriteLine Chr(18) & Chr(27) & "E" & Chr(14) & Chr(15) & PADC(Trim("No Such Serial No Created"), 30) & Chr(27) & "F"
MsgBox "no such serial no"
End If
rs1.Close
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
'Text5.SetFocus
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
Command1.Enabled = True
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

If Trim(Text13.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
Text13.SetFocus
Exit Sub
End If

If Trim(Text17.Text) = "" Then
MsgBox "Please select colliery", vbInformation, "Data Error"
Text17.SetFocus
Exit Sub
End If

If Trim(Text16.Text) = "" Then
MsgBox "Please select colliery properly", vbInformation, "Data Error"
Text17.SetFocus
Exit Sub
End If

If Val(Trim(Text13.Text)) < 5000 Then
MsgBox "RLW cannot be less than 5 Ton", vbInformation, "Data Error"
Text13.SetFocus
Exit Sub
End If
End Sub

Private Sub Command1_Click()
If camcode = 1 Then
    Command3_Click
    DoEvents
    Command3_Click
    DoEvents
    spfstwt.SetFocus
End If
Text9.Text = main.comwt.Caption
Command1.Enabled = False
End Sub


Private Sub Command1_GotFocus()
If Trim(Text10.Text) = "" Then
'Text10.SetFocus
Exit Sub
End If

If Trim(Text7.Text) = "" Then
Text7.SetFocus
Exit Sub
End If
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim ax, bx As Boolean
cctv.SetFocus
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

SavePicture Picture1.Image, "d:\wbdata\imagedata\" & Text1.Text & "_fw1.jpg"
SavePicture Picture2.Image, "d:\wbdata\imagedata\" & Text1.Text & "_fw2.jpg"
SavePicture Picture3.Image, "d:\wbdata\imagedata\" & Text1.Text & "_fw3.jpg"
SavePicture Picture4.Image, "d:\wbdata\imagedata\" & Text1.Text & "_fw4.jpg"
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
    cmdsave.Enabled = False
End If
FTP1.HostName = ips(2)
FTP1.username = unames(2)
FTP1.PassWord = passws(2)
FTP1.ChangeRemoteDir (paths(2))
If FTP1.Connect Then
Else
    MsgBox "Headquarter not connected..check connection and reload form", vbCritical
    cmdsave.Enabled = False
End If

If gpscode = 1 Then
    FTP2.HostName = "172.20.0.99"
    FTP2.username = "arsftp1"
    FTP2.PassWord = "B((1@r$f+p"
    'FTP1.ChangeRemoteDir ("root")
    If FTP2.Connect Then
    Else
        MsgBox "GPS server not connected", vbCritical
    End If
End If
tagvalid = False
Label29.Caption = ""
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
username = loginname
doorin = 0
d1 = 0
d2 = 0
Call Conn

If rfidcode = 1 Then
    Call Conn1
Else
    Text3.Locked = False
    Text4.Locked = False
    Text13.Locked = False
    Text11.Locked = False
    Text17.Locked = False
    Text17.Enabled = True
End If
unlocktxt

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

If boomcode = 1 And rfidcode = 1 Then
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

Private Sub Form_Terminate()
If rfidcode = 1 Then
If boomcode = 1 Then
    Winsock1.Close
    Winsock2.Close
End If
    main.Timer3.Enabled = True
    main.texttagg.Caption = ""
    main.tagg.Caption = ""
    validtag = False
End If
main.unloadfrm
main.invalidtags.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rfidcode = 1 Then
If boomcode = 1 Then
    Winsock1.Close
    Winsock2.Close
End If
    main.texttagg.Caption = ""
    main.tagg.Caption = ""
    main.Timer3.Enabled = True
    validtag = False
End If
main.unloadfrm
main.invalidtags.Clear
End Sub

Private Sub Addlog(ByVal titl As String, ByVal msg As String)
    'Text3.Text = Text3.Text + (Chr(13)) + Chr(10) + titl + msg
End Sub

Private Sub autoincr()
Dim pribno As Integer
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from special where date_in=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "# order by sl_no", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
    rs1.MoveLast
    pribno = Val(Right(rs1.Fields("sl_no").Value, 3))
    pribno = pribno + 1
    Text1.Text = IncrBillNo1(rs1.Fields("sl_no").Value)
Else
    pribno = 1
End If
Text1.Text = arcode + wbcode + Right(str(Year(DTPicker1.Value)), 2) + Format(Right(str(Month(DTPicker1.Value)), 2), "0#") + Format(Right(str(Day(DTPicker1.Value)), 2), "0#") + Format(pribno, "00#")
End Sub

Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = Format(Time, "HH:MM")
Text3.Text = username
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text9.Text = "0"
Text10.Text = ""
Text11.Text = ""
'Text12.Text = ""
Text13.Text = ""
Text15.Text = ""
'Text18.Text = "       "
Label20 = "0"
Label21 = "0"
Label23 = ""
DTPicker1.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
autoincr
'List1.Visible = False
sesiongen
End Sub

Private Sub List1_DblClick()
Text10.Text = tmpcode(List1.ListIndex)
custdata
'        Text5.Text = List1.List(List1.ListIndex)
'        Text6.Text = tmpcode1(List1.ListIndex)
'        Text11.Text = tmpcode2(List1.ListIndex)
'        Text12.Text = Format(tmpcode3(List1.ListIndex), "dd/mm/yyyy")
'        Text7.SetFocus
List1.Visible = False
pressf1 = False
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call List1_DblClick
End If

End Sub


Sub checkallot()
Dim tcount As Integer
dallot = 0
tcount = 0
Set rs6 = New ADODB.Recordset
rs6.Open "select count(*) from special where do_no='" + Trim(Text11.Text) + "' and date_in=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "#", co, adOpenKeyset, adLockOptimistic
tcount = rs6.Fields(0)
rs6.Close

Set rs6 = New ADODB.Recordset
rs6.Open "select * from allotment where do_no='" + Trim(Text11.Text) + "' and w_date=#" + Format(DTPicker1.Value, "mm/dd/yyyy") + "#", co, adOpenKeyset, adLockOptimistic
If rs6.RecordCount > 0 Then
    dallot = rs6.Fields("allotment").Value
    dallot = dallot - tcount
Else
    dallot = 0
End If
rs6.Close
End Sub


Private Sub List2_DblClick()
'Text17.Text = Trim(padl(List2.List(List2.ListIndex), 6))
'Text16.Text = Trim(PADR(List2.List(List2.ListIndex), 50))

Text16.Text = List2.List(List2.ListIndex)
Text17.Text = tmcode1(List2.ListIndex)

List2.Visible = False
pressf1 = False
'Text9.SetFocus
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2_DblClick
End If
End Sub

Private Sub List3_DblClick()
Text11.Text = Trim(padl(List3.List(List3.ListIndex), 11))
'checkallot
'If dallot > 0 Then
    DO_data
    List3.Visible = False
    pressf1 = False
'Else
'    MsgBox "Allotment for the day exceeded or no allotment"
'End If
End Sub

Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List3_DblClick
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

Private Sub DO_data()
If Trim(Text11.Text) <> "" Then
    Set rs3 = New ADODB.Recordset
    rs3.Open "Select * from cadata where DO_NO='" & Trim(Text11.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs3.RecordCount > 0 Then
        Text11.Text = rs3.Fields("DO_NO").Value
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))

        tmpdt.Year = Mid(rs3.Fields("do_start_date").Value, 1, 4)
        tmpdt.Month = Mid(rs3.Fields("do_start_date").Value, 5, 2)
        tmpdt.Day = Mid(rs3.Fields("do_start_date").Value, 7, 2)
        DTPicker3.Value = tmpdt.Value   '   Text12.Text = rs3.Fields("S_DATE").Value
        
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
        tmpdt.Year = Mid(rs3.Fields("do_end_date").Value, 1, 4)
        tmpdt.Month = Mid(rs3.Fields("do_end_date").Value, 5, 2)
        tmpdt.Day = Mid(rs3.Fields("do_end_date").Value, 7, 2)

        
        DTPicker4.Value = tmpdt.Value '
        stadate = DTPicker3.Value
        enddate = DTPicker4.Value
        
        Text10.Text = rs3.Fields("custcd").Value
        Text6.Text = rs3.Fields("state_code").Value
        If Trim(Text6.Text) = "" Then
            Text6.Enabled = True
        Else
            Text6.Enabled = False
        End If
        Text15.Text = " "
        Text7.Text = rs3.Fields("grade").Value
        Label20.Caption = Val(rs3.Fields("do_qty").Value) * 1000
        Text5.Text = rs3.Fields("purchaser").Value
        Text10.Text = rs3.Fields("custcd").Value
        itmdatashow
        Set rs3 = New ADODB.Recordset
        rs3.Open "Select sum(Abs(second_wt-first_wt)) from special where do_no='" & Trim(Text11.Text) & "' and second_wt>0", co, adOpenKeyset, adLockOptimistic
'        MsgBox rs3.Fields(0).Value
        If rs3.RecordCount > 0 Then
            If IsNumeric(rs3.Fields(0)) Then
                Label21.Caption = Val(Label20.Caption) - Val(rs3.Fields(0))
            Else
                Label21.Caption = Val(Label20.Caption)
            End If
        End If
        If Text6.Text = "" Then
            Text6.SetFocus
        Else
            'Text17.SetFocus
        End If
        If Val(Label21.Caption) < chkwt2 Then
            Label23.Caption = "Weighment not possible, balance quantity less than " & CStr(Abs(chkwt1 - chkwt2))
        Else
            Label23.Caption = "Balance OK Proceed with weighment"
        End If
    Else
        Text5.Text = ""
        Text6.Text = ""
        Text10.Text = ""
        Text11.Text = ""
'        Text12.Text = ""
'        Text13.Text = ""
        MsgBox "This code does not exist in Master", vbInformation, "Code Not Found"
        cmdsave.Enabled = False
        Exit Sub
    End If
Else
    Text5.Text = ""
    Text6.Text = ""
    Text10.Text = ""
    Text11.Text = ""
 '   Text12.Text = ""
 '   Text13.Text = ""
 Exit Sub
End If
If DTPicker4.Value < Date Then
    MsgBox "DO Validity Expired.."
    Text5.Text = ""
    Text6.Text = ""
    Text10.Text = ""
    Text11.Text = ""
'    Text12.Text = ""
'    Text13.Text = ""
'    Text11.SetFocus
    cmdsave.Enabled = False
    Exit Sub
End If


If timecheck(Text10.Text) = False Then
    MsgBox "Weighment not allowed..Invalid Time Range"
        Text5.Text = ""
    Text6.Text = ""
    Text10.Text = ""
    Text11.Text = ""
'    Text12.Text = ""
'    Text13.Text = ""
'    Text11.SetFocus
    cmdsave.Enabled = False
End If

End Sub

Private Sub custdata()
If Trim(Text10.Text) <> "" Then
    Set rs2 = New ADODB.Recordset
    rs2.Open "Select * from consignee where c_code='" & Trim(Text10.Text) & "'", co, adOpenKeyset, adLockOptimistic
    If rs2.RecordCount > 0 Then
        Text10.Text = rs2.Fields("C_CODE").Value
        Text5.Text = rs2.Fields("c_name").Value
    End If
End If
End Sub


Private Sub Text11_Change()
If rfidcode = 1 Then
checkallot
If dallot <= 0 Then
    MsgBox "Allotment for the day exceeded or no allotment"
    cmdsave.Enabled = False
    Command1.Enabled = False
End If

DO_data
End If
End Sub


Private Sub Text11_GotFocus()
If rfidcode = 0 Then
DO_list
'List1.Visible = False
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)
For i = 0 To List3.ListCount - 1
      If Trim(Text11.Text) = Left(List3.List(i), Len(Trim(Text11.Text))) Then
                    List3.ListIndex = i
                    Exit Sub
        End If
    Next i

End If
End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
If rfidcode = 0 Then
If KeyCode = 112 Then
    If List3.Visible = True Then
        List3.Visible = False
        pressf1 = False
    Else
        pressf1 = True
        List3.Visible = True
    End If
End If

If KeyCode = 40 Then
    List3.SetFocus
End If
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If rfidcode = 0 Then
If Trim(Text11.Text) <> "" Then
    checkallot
    If dallot <= 0 Then
        MsgBox "Allotment for the day exceeded or no allotment"
        cmdsave.Enabled = False
        Command1.Enabled = False
    End If
    DO_data
Else
    DOSHOW
End If
End If
End If
End Sub

Private Sub Text12_Change()
Text18.Text = Left(Text12.Text, 4) + "XXXXXXXXXXXXXX" + Right(Text12.Text, 4)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text11.SetFocus
End If
End Sub

Private Sub Text13_LostFocus()
If Trim(Text13.Text) = "" Then
    MsgBox "Please fill RLW"
    Text13.SetFocus
ElseIf IsNumeric(Trim(Text13.Text)) = False Then
    Text13.Text = ""
    MsgBox "Please fill numeric value for RLW"
    Text13.SetFocus
End If
End Sub





Private Sub Text17_Change()
If rfidcode = 1 Then
colldatashow
End If
End Sub


Private Sub Text17_GotFocus()
If rfidcode = 0 Then
collhlp
Text17.SelStart = 0
Text17.SelLength = Len(Text17.Text)
For i = 0 To List3.ListCount - 1
      If Trim(Text17.Text) = Left(List2.List(i), Len(Trim(Text17.Text))) Then
                    List3.ListIndex = i
                    Exit Sub
        End If
    Next i
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If rfidcode = 0 Then
If KeyAscii = 13 Then
    If Trim(Text17.Text) <> "" Then
        If pressf1 = True Then
            Text16.Text = List2.List(List2.ListIndex)
            Text17.Text = tmcode1(List2.ListIndex)
            pressf1 = False
        Else
            colldatashow
        End If
    End If
End If
End If
End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
If rfidcode = 0 Then
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
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.SetFocus
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


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text17.SetFocus
End If
End Sub

Private Sub itmdatashow()
'Set rs = New ADODB.Recordset
'rs.Open "Select * from mater where m_code='" & Trim(Text7.Text) & "'", co, adOpenKeyset, adLockOptimistic
'If rs.RecordCount > 0 Then
'rs.MoveFirst
'Text7.Text = rs.Fields("m_code").Value
'Text8.Text = rs.Fields("m_name").Value
'Else
'MsgBox "Item Not found in Item Master", vbCritical, "Item not Found "
'Text8.Text = ""
''Text7.SetFocus
'Exit Sub
'End If
End Sub

Private Sub colldatashow()
Set rs = New ADODB.Recordset
rs.Open "Select * from colliery where coll_code='" & Trim(Text17.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
Text17.Text = rs.Fields("coll_code").Value
Text16.Text = rs.Fields("coll_desc").Value
Else
MsgBox "Item Not found ", vbCritical, "Item not Found "
Text16.Text = ""
Text17.SetFocus
Exit Sub
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
'Text9.SetFocus
Else
Call KeyPress1(KeyAscii)

End If

End Sub



Private Sub Text9_Change()
'main.comwt.Caption = Text9.Text
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text9.Text) > 0 Then
cmdsave.SetFocus
Else
If Command1.Enabled = True Then Command1.SetFocus
End If
End If

End Sub


Private Sub showtagdata()
On Error GoTo errm
Timer2.Enabled = False
Conn1
Set rs4 = New ADODB.Recordset
rs4.Open "Select * from tags where tagno = '" + Trim(Texttag) + "' and valid='1'", co1, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
rs4.MoveFirst
Text12.Text = Trim(Texttag.Text)
'Command4.Enabled = False
Text14.Text = rs4.Fields("tagtrips").Value
'rs1.Fields("issue").Value = CDate(DTPicker1.Value)
'rs1.Fields("expiry").Value = CDate(DTPicker2.Value)
Text4.Text = rs4.Fields("V_no").Value
Text11.Text = rs4.Fields("do_no").Value
Text13.Text = rs4.Fields("rlw").Value
Text17.Text = rs4.Fields("coll_code").Value
Text7.Text = rs4.Fields("Tm_CODE").Value
wtmode = rs4.Fields("mode").Value
'Trim(Text10.Text) = rs1.Fields("TC_CODE").Value

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

End If
rs4.Close
Exit Sub

errm:
MsgBox Err.Description
End Sub

Sub checksock1()
If rfidcode = 1 And boomcode = 1 Then
If Winsock1.State <> "7" Then
Winsock1.Close
Winsock1.Connect
End If
End If
End Sub

Sub checksock2()
If rfidcode = 1 And boomcode = 1 Then
If Winsock2.State <> "7" Then
Winsock2.Close
Winsock2.Connect
End If
End If
End Sub



Private Sub Timer1_Timer()
'On Error Resume Next

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
    cmdsave.Enabled = False
    Command1.Enabled = False
End If
Text2.Text = Format(Time, "HH:MM")

If wtmode = "D" And rfidcode = 1 Then
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

If wtmode = "R" And rfidcode = 1 Then
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

If rfidcode = 1 Then
doorin1 = doorin
d11 = d1
d21 = d2
End If
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
Label29.Caption = "GATE-OUT CLOSING : " + CStr(15 - tctr)
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
Label29.Caption = "GATE-IN CLOSING : " + CStr(15 - tctr1)
End If
End Sub


Private Sub DO_list()
If List3.ListCount > 0 Then
    List3.Clear
    List3.Refresh
End If

Set rs4 = New ADODB.Recordset
rs4.Open "Select * from cadata order by do_start_date desc", co, adOpenKeyset, adLockOptimistic
'rs4.Open "Select * from DO_Master", co, adOpenKeyset, adLockOptimistic
If rs4.RecordCount > 0 Then
    rs4.MoveFirst
Else
    MsgBox "NO RECOERDS"
End If

g = 0
ReDim tmp(rs4.RecordCount)
While Not rs4.EOF
If Trim(Mid(rs4.Fields("do_start_date").Value, 1, 4)) <> "" Then
    List3.AddItem padl(rs4.Fields("DO_NO").Value, 11)
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
        tmpdt.Year = Mid(rs4.Fields("do_start_date").Value, 1, 4)
        tmpdt.Month = Mid(rs4.Fields("do_start_date").Value, 5, 2)
        tmpdt.Day = Mid(rs4.Fields("do_start_date").Value, 7, 2)
        tmp(g) = tmpdt.Value
    End If
    g = g + 1
    rs4.MoveNext
Wend
End Sub

Private Sub DOSHOW()
If List3.ListCount > 0 Then
    List3.Clear
    List3.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from cadata where custcd='" & Trim(Text10.Text) & "'", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    rs.MoveFirst
End If

g = 0
ReDim tmp(rs.RecordCount)
While Not rs.EOF
    List3.AddItem padl(rs.Fields("do_no").Value, 11)
        tmpdt.Value = CDate(Format("01/01/2000", "dd/mm/yyyy"))
        tmpdt.Year = Mid(rs.Fields("do_start_date").Value, 1, 4)
        tmpdt.Month = Mid(rs.Fields("do_start_date").Value, 5, 2)
        tmpdt.Day = Mid(rs.Fields("do_start_date").Value, 7, 2)
    tmp(g) = tmpdt.Value
    g = g + 1
    rs.MoveNext
Wend


End Sub

Private Sub collhlp()
Set rs5 = New ADODB.Recordset
'rs5.Open "Select * from colliery order by coll_des", co, adOpenKeyset, adLockOptimistic
rs5.Open "select * from colliery", co, adOpenKeyset, adLockOptimistic
If rs5.RecordCount > 0 Then
rs5.MoveFirst
End If
g = 0
ReDim tmcode1(rs5.RecordCount)
While Not rs5.EOF
List2.AddItem rs5.Fields("coll_desc").Value
tmcode1(g) = rs5.Fields("coll_code").Value
g = g + 1
rs5.MoveNext
Wend
End Sub
