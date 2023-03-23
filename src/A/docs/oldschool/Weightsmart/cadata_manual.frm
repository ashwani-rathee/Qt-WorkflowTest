VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ca_manual 
   Caption         =   "CA Data Manual"
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
      Height          =   6135
      Left            =   12240
      TabIndex        =   56
      Top             =   960
      Width           =   2655
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
      Left            =   13680
      TabIndex        =   55
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   15000
      TabIndex        =   50
      Top             =   1680
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
         TabIndex        =   51
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
         TabIndex        =   52
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   240
      TabIndex        =   49
      Top             =   720
      Width           =   14775
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
         ItemData        =   "cadata_manual.frx":0000
         Left            =   2040
         List            =   "cadata_manual.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Payment Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   4920
         TabIndex        =   99
         Top             =   240
         Width           =   3375
         Begin VB.TextBox Text45 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   28
            Top             =   6480
            Width           =   1935
         End
         Begin VB.TextBox Text44 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   27
            Top             =   6120
            Width           =   1695
         End
         Begin VB.TextBox Text43 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   26
            Top             =   5760
            Width           =   1695
         End
         Begin VB.TextBox Text42 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   25
            Top             =   5400
            Width           =   1695
         End
         Begin VB.TextBox Text41 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   24
            Top             =   4800
            Width           =   1935
         End
         Begin VB.TextBox Text40 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   23
            Top             =   4440
            Width           =   1695
         End
         Begin VB.TextBox Text38 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   21
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox Text37 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   20
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox Text36 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   19
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text34 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   17
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox Text33 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   16
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox Text32 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Text30 
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
            Left            =   1320
            MaxLength       =   11
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker Text31 
            Height          =   375
            Left            =   1320
            TabIndex        =   14
            Top             =   720
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin MSComCtl2.DTPicker Text35 
            Height          =   375
            Left            =   1320
            TabIndex        =   18
            Top             =   2400
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin MSComCtl2.DTPicker Text39 
            Height          =   375
            Left            =   1320
            TabIndex        =   22
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Details"
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
            Left            =   120
            TabIndex        =   116
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tax %"
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
            TabIndex        =   115
            Top             =   6480
            Width           =   1215
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tax Type"
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
            TabIndex        =   114
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Qty Bal"
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
            TabIndex        =   113
            Top             =   5760
            Width           =   1215
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Amt Bal"
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
            TabIndex        =   112
            Top             =   5400
            Width           =   1215
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank3"
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
            TabIndex        =   111
            Top             =   4800
            Width           =   1215
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Amt3"
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
            TabIndex        =   110
            Top             =   4440
            Width           =   1215
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Dt3"
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
            TabIndex        =   109
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft No3"
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
            TabIndex        =   108
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank2"
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
            TabIndex        =   107
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Amt2"
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
            TabIndex        =   106
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Dt2"
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
            TabIndex        =   105
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft No2"
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
            TabIndex        =   104
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank1"
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
            TabIndex        =   103
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Amt1"
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
            TabIndex        =   102
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft Dt1"
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
            TabIndex        =   101
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Draft No1"
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
            TabIndex        =   100
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CA Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   8520
         TabIndex        =   80
         Top             =   120
         Width           =   3255
         Begin VB.TextBox Text29 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   46
            Top             =   6480
            Width           =   1215
         End
         Begin VB.TextBox Text28 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   45
            Top             =   6120
            Width           =   1215
         End
         Begin VB.TextBox Text27 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   44
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox Text26 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   43
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox Text25 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   42
            Top             =   5040
            Width           =   1215
         End
         Begin VB.TextBox Text24 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   41
            Top             =   4680
            Width           =   1215
         End
         Begin VB.TextBox Text23 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   40
            Top             =   4320
            Width           =   1215
         End
         Begin VB.TextBox Text22 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   39
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox Text21 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   38
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox Text20 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   37
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Text19 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   36
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox Text18 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   35
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox Text17 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   34
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox Text16 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text15 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   32
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text14 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   31
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Text13 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   30
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text12 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " High Edu Rate"
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
            TabIndex        =   98
            Top             =   6480
            Width           =   1815
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Edu Cess Rate"
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
            TabIndex        =   97
            Top             =   6120
            Width           =   1815
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cent Exc Cess"
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
            TabIndex        =   96
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SPROV Chrg"
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
            TabIndex        =   95
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bazar Fee"
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
            TabIndex        =   94
            Top             =   5040
            Width           =   1815
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " WRC"
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
            TabIndex        =   93
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SLC"
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
            TabIndex        =   92
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Wmnt. Charge"
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
            TabIndex        =   91
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " C.E. Cess"
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
            TabIndex        =   90
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SED"
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
            TabIndex        =   89
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Royalty"
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
            TabIndex        =   88
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Basic Rate"
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
            TabIndex        =   87
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CST No"
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
            TabIndex        =   86
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VAT Tin No"
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
            TabIndex        =   85
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Commissionerate"
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
            TabIndex        =   84
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Division"
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
            TabIndex        =   83
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Range"
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
            TabIndex        =   82
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Exc Reg No"
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
            TabIndex        =   81
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Customer Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   73
         Top             =   4200
         Width           =   4575
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
            Left            =   1920
            TabIndex        =   12
            Top             =   1800
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
            Left            =   1920
            TabIndex        =   11
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text46 
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
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   10
            Top             =   1080
            Width           =   2535
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
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   9
            Top             =   720
            Width           =   2535
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   8
            Top             =   360
            Width           =   1455
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
            Left            =   120
            TabIndex        =   78
            Top             =   1800
            Width           =   1860
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " State"
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
            TabIndex        =   74
            Top             =   1440
            Width           =   1845
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Address"
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
            TabIndex        =   117
            Top             =   1080
            Width           =   1815
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
            Left            =   3480
            TabIndex        =   77
            Top             =   405
            Visible         =   0   'False
            Width           =   615
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
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   1860
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
            Left            =   120
            TabIndex        =   75
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "DO Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   4575
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
            Left            =   1920
            TabIndex        =   47
            Top             =   3240
            Width           =   1695
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
            Left            =   1920
            TabIndex        =   7
            Top             =   2880
            Width           =   1695
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   3
            Top             =   1440
            Width           =   2055
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   1
            Top             =   720
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   2520
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   2160
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin VB.TextBox Text11 
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
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   0
            Top             =   360
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1920
            TabIndex        =   2
            Top             =   1080
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
            Format          =   90439681
            CurrentDate     =   41137
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1920
            TabIndex        =   4
            Top             =   1800
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
            Format          =   90439681
            CurrentDate     =   41137
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
            Left            =   120
            TabIndex        =   72
            Top             =   2160
            Width           =   1815
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
            Left            =   120
            TabIndex        =   71
            Top             =   2520
            Width           =   1815
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
            Left            =   3720
            TabIndex        =   70
            Top             =   3000
            Width           =   735
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
            Left            =   3720
            TabIndex        =   69
            Top             =   3360
            Width           =   735
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
            Left            =   120
            TabIndex        =   68
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Application Date"
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
            TabIndex        =   67
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Application No"
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
            TabIndex        =   66
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DO Date"
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
            TabIndex        =   65
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "11 digit  required"
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   3960
            TabIndex        =   64
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DO Number"
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
            TabIndex        =   63
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Unit Number"
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
            TabIndex        =   79
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame editfrm 
         Height          =   735
         Left            =   12000
         TabIndex        =   57
         Top             =   6480
         Width           =   2655
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
            Left            =   1440
            TabIndex        =   61
            Top             =   240
            Width           =   1095
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
            TabIndex        =   60
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
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1215
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
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
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
         TabIndex        =   119
         Top             =   6840
         Width           =   1815
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
      Left            =   14520
      TabIndex        =   48
      ToolTipText     =   "Unload Form"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CA Data Manual"
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
      Left            =   5760
      TabIndex        =   53
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   240
      TabIndex        =   54
      Top             =   120
      Width           =   14775
   End
End
Attribute VB_Name = "ca_manual"
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
Dim cafound As Boolean
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
              
        cafound = False
        Set rs3 = New ADODB.Recordset
        rs3.Open "select * from cadata where do_no='" & Trim(Text1.Text) & "'", co, adOpenKeyset, adLockOptimistic
        If rs3.RecordCount > 0 Then
            rs3.MoveFirst
            cafound = True
        Else
            rs3.AddNew
            cafound = False
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
            MsgBox "Invalid product code, please enter again"
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
            
            rs3.Fields("unit") = Trim(Text11.Text)
            rs3.Fields("purchaser") = Trim(Text3.Text)
            rs3.Fields("destination") = Trim(Text46.Text)
            rs3.Fields("state_code") = Trim(Text5.Text)
            rs3.Fields("grade") = Trim(Text7.Text)
            rs3.Fields("do_no") = Trim(Text1.Text)
            rs3.Fields("do_date") = Format(DTPicker3, "0#") + Format(DTPicker3.Month, "0#") + Format(DTPicker3.Year, "000#")
            rs3.Fields("appl_no") = Trim(Text9.Text)
            rs3.Fields("appl_date") = Format(DTPicker4, "0#") + Format(DTPicker4.Month, "0#") + Format(DTPicker4.Year, "000#")
            rs3.Fields("do_qty") = Val(Trim(Text6.Text))
            rs3.Fields("draft_no1") = Trim(Text30.Text)
            rs3.Fields("draft_dt1") = Format(Text31.Day, "0#") + Format(Text31.Month, "0#") + Format(Text31.Year, "000#")
            rs3.Fields("draft_amt1") = Val(Trim(Text32.Text))
            rs3.Fields("bank1") = Trim(Text33.Text)
            rs3.Fields("draft_no2") = Trim(Text34.Text)
            rs3.Fields("draft_dt2") = Format(Text35.Day, "0#") + Format(Text35.Month, "0#") + Format(Text35.Year, "000#")
            rs3.Fields("draft_amt2") = Val(Trim(Text36.Text))
            rs3.Fields("bank2") = Trim(Text37.Text)
            rs3.Fields("draft_no3") = Trim(Text38.Text)
            rs3.Fields("draft_dt3") = Format(Text39.Day, "0#") + Format(Text39.Month, "0#") + Format(Text39.Year, "000#")
            rs3.Fields("draft_amt3") = Val(Trim(Text40.Text))
            rs3.Fields("bank3") = Trim(Text41.Text)
            rs3.Fields("amtbalance") = Val(Trim(Text42.Text))
            rs3.Fields("qtybalance") = Val(Trim(Text43.Text))
            rs3.Fields("taxtype") = Trim(Text144.Text)
            rs3.Fields("tax_percent") = Val(Trim(Text45.Text))
            rs3.Fields("custcd") = Trim(Text2.Text)
            rs3.Fields("exc_reg_no") = Trim(Text12.Text)
            rs3.Fields("range") = Trim(Text13.Text)
            rs3.Fields("division") = Trim(Text14.Text)
            rs3.Fields("commissionerate") = Trim(Text15.Text)
            rs3.Fields("vat_tin_no") = Trim(Text16.Text)
            rs3.Fields("cst_no") = Trim(Text17.Text)
            rs3.Fields("basic_rate") = Val(Trim(Text18.Text))
            rs3.Fields("royalty") = Val(Trim(Text19.Text))
            rs3.Fields("sed") = Val(Trim(Text20.Text))
            rs3.Fields("clean_engy_cess") = Val(Trim(Text21.Text))
            rs3.Fields("weighment_chg") = Val(Trim(Text22.Text))
            rs3.Fields("slc") = Val(Trim(Text23.Text))
            rs3.Fields("wrc") = Val(Trim(Text24.Text))
            rs3.Fields("bazar_fee") = Val(Trim(Text25.Text))
            rs3.Fields("sprovchrg") = Val(Trim(Text26.Text))
            rs3.Fields("cent_exc_rate") = Val(Trim(Text27.Text))
            rs3.Fields("edu_cess_rate") = Val(Trim(Text28.Text))
            rs3.Fields("high_edu_rate") = Val(Trim(Text29.Text))
            rs3.Fields("do_start_date") = Format(DTPicker1, "0#") + Format(DTPicker1.Month, "0#") + Format(DTPicker1.Year, "000#")
            rs3.Fields("do_end_date") = Format(DTPicker2, "0#") + Format(DTPicker2.Month, "0#") + Format(DTPicker2.Year, "000#")
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
Text11.Text = Format(arcode, "0#") + Format(wbcode, "0#")
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

