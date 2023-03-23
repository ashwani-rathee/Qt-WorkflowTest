VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form tagrchange 
   Caption         =   "TAG ISSUE"
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   13455
      Begin VB.CommandButton Command3 
         Caption         =   "CHANGE BEFORE 1ST WT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   9000
         TabIndex        =   83
         Top             =   6240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text12 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   300
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox Text14 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox Texttime 
         Enabled         =   0   'False
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
         Left            =   11640
         TabIndex        =   79
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text9 
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
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   6360
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   11880
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   48
         Top             =   6540
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
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
         ItemData        =   "tagchange.frx":0000
         Left            =   2280
         List            =   "tagchange.frx":0002
         TabIndex        =   47
         Top             =   1560
         Width           =   1695
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
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   14
         Top             =   6180
         Width           =   1170
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
         Left            =   9000
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text28 
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
         Enabled         =   0   'False
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
         Left            =   9960
         MaxLength       =   149
         TabIndex        =   43
         Top             =   5640
         Width           =   3210
      End
      Begin VB.TextBox Text27 
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
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   41
         Top             =   4320
         Width           =   3210
      End
      Begin VB.TextBox Text26 
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
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   39
         Top             =   3960
         Width           =   3210
      End
      Begin VB.TextBox Text25 
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
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   37
         Top             =   3600
         Width           =   3210
      End
      Begin VB.TextBox Text24 
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
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   35
         Top             =   3120
         Width           =   3210
      End
      Begin VB.TextBox Text23 
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
         Enabled         =   0   'False
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2760
         Width           =   3210
      End
      Begin VB.TextBox Text22 
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
         Enabled         =   0   'False
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2400
         Width           =   3210
      End
      Begin VB.TextBox Text21 
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
         Enabled         =   0   'False
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2040
         Width           =   3210
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CHANGE BEFORE 2ND WT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   11160
         TabIndex        =   23
         Top             =   6240
         Width           =   2055
      End
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
         Left            =   5520
         TabIndex        =   10
         Top             =   4920
         Visible         =   0   'False
         Width           =   3015
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
         Left            =   4320
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   2280
         TabIndex        =   12
         Top             =   5280
         Width           =   3255
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
         Left            =   2280
         TabIndex        =   11
         Top             =   4920
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   9360
         TabIndex        =   7
         Top             =   1440
         Width           =   3900
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
            Left            =   2040
            TabIndex        =   9
            Top             =   120
            Width           =   1725
         End
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
            TabIndex        =   8
            Top             =   120
            Width           =   1935
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   10200
         TabIndex        =   25
         Top             =   840
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
         Format          =   73596929
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73596929
         CurrentDate     =   39961
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   45
         Text            =   "1"
         Top             =   5820
         Width           =   615
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   6240
         TabIndex        =   53
         Top             =   1620
         Width           =   3495
         Begin VB.OptionButton Option2 
            Caption         =   "RECEIVE"
            Enabled         =   0   'False
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
            Left            =   1680
            TabIndex        =   55
            Top             =   60
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "DISPATCH"
            Enabled         =   0   'False
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
            Left            =   0
            TabIndex        =   54
            Top             =   60
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   6240
         TabIndex        =   57
         Top             =   960
         Width           =   3495
         Begin VB.OptionButton Option3 
            Caption         =   "SPECIAL"
            Enabled         =   0   'False
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
            Left            =   0
            TabIndex        =   59
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "SIMPLE"
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
            Left            =   1680
            TabIndex        =   58
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WEIGHMENT TYPE"
            Enabled         =   0   'False
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
            Left            =   480
            TabIndex        =   60
            Top             =   60
            Width           =   2070
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   6360
         TabIndex        =   61
         Top             =   2160
         Width           =   6855
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
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   2775
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
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2400
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
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1320
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
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   960
            Width           =   4815
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
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   600
            Width           =   1455
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
            Left            =   1800
            MaxLength       =   11
            TabIndex        =   63
            Top             =   240
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1800
            TabIndex        =   68
            Top             =   2040
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
            Format          =   73596929
            CurrentDate     =   40669
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1800
            TabIndex        =   69
            Top             =   1680
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
            Format          =   73596929
            CurrentDate     =   40669
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   76
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   75
            Top             =   1800
            Width           =   1050
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   74
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   73
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   72
            Top             =   2520
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   71
            Top             =   1440
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   70
            Top             =   1080
            Width           =   630
         End
      End
      Begin VB.TextBox Text8 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "OLD TAG"
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
         Left            =   720
         TabIndex        =   82
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         TabIndex        =   78
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label35 
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
         Left            =   3480
         TabIndex        =   56
         Top             =   6240
         Width           =   285
      End
      Begin VB.Label status 
         BackColor       =   &H000000C0&
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
         ForeColor       =   &H000000C0&
         Height          =   795
         Left            =   5760
         TabIndex        =   52
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Code"
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
         Left            =   720
         TabIndex        =   49
         Top             =   6660
         Width           =   1035
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle Type"
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
         Left            =   720
         TabIndex        =   46
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Driver Pic"
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
         Left            =   8760
         TabIndex        =   44
         Top             =   5760
         Width           =   1035
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Driver Phone"
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
         Left            =   720
         TabIndex        =   42
         Top             =   4440
         Width           =   1365
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Driver Address"
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
         Left            =   720
         TabIndex        =   40
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Truck Driver"
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
         Left            =   720
         TabIndex        =   38
         Top             =   3720
         Width           =   1290
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Owner Phone"
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
         Left            =   720
         TabIndex        =   36
         Top             =   3240
         Width           =   1380
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Owner Email"
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
         Left            =   720
         TabIndex        =   34
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Owner Address"
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
         Left            =   720
         TabIndex        =   32
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Truck Owner"
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
         Left            =   720
         TabIndex        =   30
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Expiry"
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
         Left            =   720
         TabIndex        =   28
         Top             =   1260
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "No of Trips"
         Enabled         =   0   'False
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
         Left            =   720
         TabIndex        =   26
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "NEW TAG"
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
         Left            =   720
         TabIndex        =   22
         Top             =   780
         Width           =   1125
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
         Left            =   7440
         TabIndex        =   20
         Top             =   360
         Width           =   1155
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
         Left            =   720
         TabIndex        =   19
         Top             =   6300
         Width           =   510
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
         Left            =   4320
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   720
         TabIndex        =   17
         Top             =   5400
         Width           =   810
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
         Left            =   720
         TabIndex        =   16
         Top             =   5040
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   14880
      TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   240
         Width           =   3255
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
      Left            =   13680
      TabIndex        =   0
      ToolTipText     =   "Unload Form"
      Top             =   600
      Width           =   375
   End
   Begin MSComCtl2.DTPicker tmpdt 
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
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
      Format          =   73596929
      CurrentDate     =   40669
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE TAG FOR ACTIVE WEIGHMENT"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   4680
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   675
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   13455
   End
End
Attribute VB_Name = "tagrchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Dim num As Integer
Dim recfound As Boolean
Dim pressf1 As Boolean
Dim g As Integer
Dim chkpass As Integer
Dim stadate As Date
Dim enddate As Date
Dim dallot As Integer

Private Sub getTag()
Dim OutData(30000) As Byte
Dim TagCounter(20) As Byte
Dim str As String

a = Inventory(1, OutData(0), TagCounter(0))

TagNum = TagCounter(0)
If TagNum > 0 Then
    a = CleanInventory()
    For j = 0 To 0
        For i = 0 To OutData(30 * j + 6) - 1
        str = str + Right$("00" & Hex$(OutData(30 * j + 7 + i)), 2)
        Next i
    Next j
    
    Text1.Text = str
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
            
            ret = rdy_read(3, 0, 8, 0, TagCount(0), DataLen(0), Data(0), ReadLen(0), AntID(0), ReadCount(0))
            If ret <> 0 Then
                Text1.Text = ""
                Exit Sub
            End If
            
            For j = 2 To 13
                str = str + Right$("00" & Hex$(Data(j)), 2)
            Next
If str <> Label1.Caption Then
    Text1.Text = str
End If
End Sub


Private Sub Command1_Click()
Dim costr As String
Dim wmode As String
Dim wtype As String

If Trim(Text9.Text) = "" Then
    MsgBox "Enter Remarks"
    Exit Sub
End If

b = MsgBox("Sure? Change Tag ? ", vbOKCancel, "Print ?")
If b = vbCancel Then
    MsgBox "Saving Cancelled"
    Exit Sub
End If

If Option2.Value = True Then
wmode = "R"
Else
wmode = "D"
End If

If Option3.Value = True Then
wtype = "1"
Else
wtype = "2"
End If


On Error GoTo errmess
costr = "update tags set valid='0'" _
+ " where tagno = '" + Trim(Text12.Text) + "' and valid='2'"
co1.Execute costr

costr = "update special set tag='" + Text1.Text _
+ "' where v_no = '" + Trim(Text4.Text) + "' and second_wt=0"
co1.Execute costr

costr = "insert into tags(tagno,tagtrips,issue,expiry,tc_code,v_no,tm_CODE,rlw,do_no,coll_code,valid," _
+ "owner,owner_address,owner_email,owner_phone,driver,driver_address,driver_phone,photo,trips_done,v_type,unit,mode,wmode)" _
+ " values ('" + Text1.Text + "','" + Trim(Text2.Text) + "','" + Format(Year(DTPicker1.Value), "00##") _
+ Format(Month(DTPicker1.Value), "0#") + Format(Day(DTPicker1.Value), "0#") + "','" + Format(Year(DTPicker2.Value), "00##") _
+ Format(Month(DTPicker2.Value), "0#") + Format(Day(DTPicker2.Value), "0#") + "','" + Trim(Text10.Text) + "','" _
+ Trim(Text4.Text) + "','" + Trim(Text7.Text) + "','" + Trim(Text13.Text) + "','" + Trim(Text11.Text) + "','" _
+ Trim(Text17.Text) + "','2','" + Trim(Text21.Text) + "','" + Trim(Text22.Text) + "','" + Trim(Text23.Text) + "','" + Trim(Text24.Text) _
+ "','" + Trim(Text25.Text) + "','" + Trim(Text26.Text) + "','" + Trim(Text27.Text) + "','" + Trim(Text28.Text) _
+ "',0,'" + Combo1.Text + "','" + Trim(Text3.Text) + "','" + Trim(wmode) + "','" + Trim(wtype) + "')"
'MsgBox costr
co1.Execute costr

co1.Execute "insert into tagexceptions (edate,etime,tagno,v_no,ccode,sl_no,type,description,o_name,area,wb)" _
+ " values ('" + Format(DTPicker1.Value, "yyyy/mm/dd") + "','" + Trim(Texttime.Text) + "','" + Trim(Text1.Text) _
+ "','" + Trim(Text4.Text) + "','" + Trim(Text10.Text) + "','" + "" + "','" + "TAG CHANGE" + "','" _
+ Trim(Text9.Text) + "','" + loginname + "','" + Left(Text3.Text, 2) + "','" + Right(Text3.Text, 2) + "' );"


Set rs5 = New ADODB.Recordset
rs5.Open "Select * from special where tag='" & Trim(Text12.Text) & "' and second_wt=0", co, adOpenKeyset, adLockOptimistic
If rs5.RecordCount > 0 Then
rs5.MoveFirst
While rs5.EOF = False
    rs5.Fields("tag") = Trim(Text1.Text)
    rs5.Update
    rs5.MoveNext
Wend
End If
rs5.Close

MsgBox "Tag information saved"
b = 1
    While b = 1
    b = MsgBox("Print Tag ", vbOKCancel, "Print ?")
    If b = 1 Then
       printtag
    End If
    Wend
Exit Sub

errmess:
MsgBox Err.Description
End Sub

Private Sub Command1_GotFocus()
If Trim(Text1.Text) = "" Then
MsgBox "Invalid Tag Number. Please scan a valid Tag", vbInformation, "Data Error"
'Text4.SetFocus
Exit Sub
End If

If Trim(Combo1.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Combo1.SetFocus
Exit Sub
End If

If Trim(Text4.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text4.SetFocus
Exit Sub
End If

If Val(Trim(Text13.Text)) < 5000 Then
MsgBox "RLW cannot be less than 5 Ton", vbInformation, "Data Error"
'Text13.SetFocus
Exit Sub
End If

If Trim(Text17.Text) = "" Then
MsgBox "Please select colliery", vbInformation, "Data Error"
'Text17.SetFocus
Exit Sub
End If

If Trim(Text16.Text) = "" Then
MsgBox "Please select colliery properly", vbInformation, "Data Error"
'Text17.SetFocus
Exit Sub
End If

If Trim(Text11.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text11.SetFocus
Exit Sub
End If

If Trim(Text5.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text11.SetFocus
Exit Sub
End If

If Trim(Text6.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text11.SetFocus
Exit Sub
End If

If Trim(Text7.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text11.SetFocus
Exit Sub
End If

If Trim(Text10.Text) = "" Then
MsgBox "Please Data Fill Properly", vbInformation, "Data Error"
'Text11.SetFocus
Exit Sub
End If


End Sub

Private Sub Command2_Click()
On Error GoTo errmess
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from  tags where v_no = '" + Trim(Text4.Text) + "' order by sno desc", co1, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
    If rs1.Fields("valid") = 0 Then
        status.Caption = "INACTIVE TAG, CAN NOT CHANGE"
        Command1.Visible = False
        Command3.Visible = False
    ElseIf rs1.Fields("valid") = 1 Then
        status.Caption = "1st WT DUE..CAN BE CHANGED"
        Command1.Visible = False
        Command3.Visible = True
    Else
        status.Caption = "2nd WT DUE..CAN BE CHANGED"
        Command1.Visible = True
        Command3.Visible = False
    End If
    
    Text12.Text = rs1.Fields("tagno")
    Text2.Text = rs1.Fields("tagtrips")
    DTPicker1.Value = CDate(rs1.Fields("issue"))
    DTPicker2.Value = CDate(rs1.Fields("expiry"))
    Text10.Text = rs1.Fields("tc_code")
    Text4.Text = rs1.Fields("v_no")
    Text7.Text = rs1.Fields("tm_CODE")
    Text13.Text = rs1.Fields("rlw")
    Text11.Text = rs1.Fields("do_no")
    Text17.Text = rs1.Fields("coll_code")
    Text21.Text = rs1.Fields("owner")
    Text22.Text = rs1.Fields("owner_address")
    Text23.Text = rs1.Fields("owner_email")
    Text24.Text = rs1.Fields("owner_phone")
    Text25.Text = rs1.Fields("driver")
    Text26.Text = rs1.Fields("driver_address")
    Text27.Text = rs1.Fields("driver_phone")
    Text28.Text = rs1.Fields("photo")
    Combo1.Enabled = True
    Combo1.Text = rs1.Fields("v_type")
    Combo1.Enabled = False
    Text3.Text = rs1.Fields("unit")
    If rs1.Fields("mode") = "R" Then
        Option2.Value = True
    Else
        Option1.Value = True
    End If
    If rs1.Fields("wmode") = "2" Then
        Option4.Value = True
    Else
        Option3.Value = True
    End If
hideshow
End If
rs1.Close

'checkallot
'If dallot > 0 Then
    DO_data
    pressf1 = False
'Else
'    MsgBox "Allotment for the day exceeded or no allotment"
'End If

colldatashow
Exit Sub

errmess:
MsgBox Err.Description
End Sub

Private Sub Command3_Click()
Dim costr As String
Dim wmode As String
Dim wtype As String

If Trim(Text9.Text) = "" Then
    MsgBox "Enter Remarks"
    Exit Sub
End If

b = MsgBox("Sure? Change Tag ? ", vbOKCancel, "Print ?")
If b = vbCancel Then
    MsgBox "Saving Cancelled"
    Exit Sub
End If

If Option2.Value = True Then
wmode = "R"
Else
wmode = "D"
End If

If Option3.Value = True Then
wtype = "1"
Else
wtype = "2"
End If


On Error GoTo errmess
costr = "update tags set valid='0'" _
+ " where tagno = '" + Trim(Text12.Text) + "' and valid='1'"
co1.Execute costr

costr = "insert into tags(tagno,tagtrips,issue,expiry,tc_code,v_no,tm_CODE,rlw,do_no,coll_code,valid," _
+ "owner,owner_address,owner_email,owner_phone,driver,driver_address,driver_phone,photo,trips_done,v_type,unit,mode,wmode)" _
+ " values ('" + Text1.Text + "','" + Trim(Text2.Text) + "','" + Format(Year(DTPicker1.Value), "00##") _
+ Format(Month(DTPicker1.Value), "0#") + Format(Day(DTPicker1.Value), "0#") + "','" + Format(Year(DTPicker2.Value), "00##") _
+ Format(Month(DTPicker2.Value), "0#") + Format(Day(DTPicker2.Value), "0#") + "','" + Trim(Text10.Text) + "','" _
+ Trim(Text4.Text) + "','" + Trim(Text7.Text) + "','" + Trim(Text13.Text) + "','" + Trim(Text11.Text) + "','" _
+ Trim(Text17.Text) + "','1','" + Trim(Text21.Text) + "','" + Trim(Text22.Text) + "','" + Trim(Text23.Text) + "','" + Trim(Text24.Text) _
+ "','" + Trim(Text25.Text) + "','" + Trim(Text26.Text) + "','" + Trim(Text27.Text) + "','" + Trim(Text28.Text) _
+ "',0,'" + Combo1.Text + "','" + Trim(Text3.Text) + "','" + Trim(wmode) + "','" + Trim(wtype) + "')"
'MsgBox costr
co1.Execute costr

co1.Execute "insert into tagexceptions (edate,etime,tagno,v_no,ccode,sl_no,type,description,o_name,area,wb)" _
+ " values ('" + Format(DTPicker1.Value, "yyyy/mm/dd") + "','" + Trim(Texttime.Text) + "','" + Trim(Text1.Text) _
+ "','" + Trim(Text4.Text) + "','" + Trim(Text10.Text) + "','" + "" + "','" + "TAG CHANGE" + "','" _
+ Trim(Text9.Text) + "','" + loginname + "','" + Left(Text3.Text, 2) + "','" + Right(Text3.Text, 2) + "' );"

MsgBox "Tag information saved"
b = 1
    While b = 1
    b = MsgBox("Print Tag ", vbOKCancel, "Print ?")
    If b = 1 Then
       printtag
    End If
    Wend
Exit Sub

errmess:
MsgBox Err.Description
End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub Form_Load()
Me.Picture = main.Picture
DTPicker1.Value = Date
DTPicker2.Value = Date
'Combo1.ListIndex = 0
Texttime.Text = Format(Time, "hh:mm")
Conn1
End Sub

Function trimdate(tda As Date) As String
trimdate = Format(Day(tda), "0#") & Format(Month(tda), "0#") & Format(Year(tda), "0###")
End Function

Sub printtag()
Dim ordst As String
Dim orden As String
On Error Resume Next
tagprint.Sections(1).Controls("Label1").Caption = PADC(Trim(main.Label1.Caption), 27)
tagprint.Sections(1).Controls("Label2").Caption = PADC(Trim(main.Label2.Caption), 27)
tagprint.Sections(1).Controls("Label3").Caption = PADC(Trim(main.Label3.Caption), 27)

If Option3.Value = True Then
    tagprint.Sections(1).Controls("Label4").Caption = "Tag No : " + Text8.Text + " (Special)"
Else
    tagprint.Sections(1).Controls("Label4").Caption = "Tag No : " + Text8.Text + " (Simple)"
End If

Set rs1 = New ADODB.Recordset
rs1.Open "SELECT * from tags where tagno='" & Trim(Text1.Text) & "'", co1, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Set tagprint.DataSource = rs1
tagprint.Sections(1).Controls("Text1").DataField = "tagno"
tagprint.Sections(1).Controls("Label5").Caption = "Issue Date : " + CStr(DTPicker1.Value) + "   Expiry Date: " + CStr(DTPicker2.Value)
tagprint.Sections(1).Controls("Label6").Caption = "Vehicle Type : " + Combo1.Text + "   Vehicle No : " + Text4.Text
tagprint.Sections(1).Controls("Label7").Caption = "RLW : " + Text13.Text
tagprint.Sections(1).Controls("Label8").Caption = "Owner : " + Text21.Text
tagprint.Sections(1).Controls("Label9").Caption = "Owner Address : " + Text22.Text
tagprint.Sections(1).Controls("Label10").Caption = "Owner Email : " + Text23.Text
tagprint.Sections(1).Controls("Label11").Caption = "Owner Phone : " + Text24.Text

tagprint.Sections(1).Controls("Label12").Caption = "Driver : " + Text25.Text
tagprint.Sections(1).Controls("Label13").Caption = "Driver Address : " + Text26.Text
tagprint.Sections(1).Controls("Label14").Caption = "Driver Phone : " + Text27.Text
tagprint.Sections(1).Controls("Label15").Caption = "Driver : " + Text25.Text

tagprint.Sections(1).Controls("Label16").Caption = "Colliery : " + Text17.Text + " - " + Text16.Text
tagprint.Sections(1).Controls("Label17").Caption = "Unit Code : " + Text3.Text
tagprint.Sections(1).Controls("Label18").Caption = "DO Number : " + Text11.Text
tagprint.Sections(1).Controls("Label19").Caption = "DO Start : " + CStr(DTPicker3.Value) + "   DO End: " + CStr(DTPicker4.Value)
tagprint.Sections(1).Controls("Label20").Caption = "Party : " + Text10.Text + " - " + Text5.Text
tagprint.Sections(1).Controls("Label21").Caption = "Destination : " + Text6.Text
tagprint.Sections(1).Controls("Label22").Caption = "Material : " + Text7.Text

tagprint.Sections(1).Controls("Label23").Caption = "Ordered Quantity : " + Label20.Caption
tagprint.Sections(1).Controls("Label24").Caption = ""
tagprint.Sections(1).Controls("Label25").Caption = ""

tagprint.Sections(1).Controls("Label26").Caption = PADR(COMPAUTH, 20) & padl(" ", 2) & padl("Operator", 30)
Else
tagprint.Sections(1).Controls("Label5").Caption = PADC(Trim("No Such Tag No Created"), 30)
End If
tagprint.Show vbModal
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


Private Sub unlocktxt()
Text1.Text = ""
Text2.Text = Format(Time, "HH:MM")
Text3.Text = ""
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
DTPicker1.Value = Date
DTPicker3.Value = Date
DTPicker4.Value = Date
'autoincr
'List1.Visible = False
'List2.Visible = False
List3.Visible = False
'sesiongen
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

Sub hideshow()
If Option4.Value = True Then
    Frame6.Visible = False
    Text16.Visible = False
    Text17.Visible = False
    List2.Visible = False
    Text26.Visible = False
    Text27.Visible = False
Else
    Frame6.Visible = True
    Text16.Visible = True
    Text17.Visible = True
    List2.Visible = True
    Text26.Visible = True
    Text27.Visible = True
End If
End Sub

Private Sub Option3_Click()
hideshow
End Sub

Private Sub Option4_Click()
hideshow
End Sub

Private Sub Text1_Change()
Text8.Text = Left(Text1.Text, 4) + "XXXXXXXXXXXXXXXXXX" + Right(Text1.Text, 4)
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Text2.SetFocus
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

        
        DTPicker4.Value = tmpdt.Value '    Text14.Text = rs3.Fields("end_DATE").Value
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
        If Text6.Text = "" Then
            'Text6.SetFocus
        Else
            'Text17.SetFocus
        End If
    Else
        Text5.Text = ""
        Text6.Text = ""
        Text10.Text = ""
        Text11.Text = ""
'        Text12.Text = ""
'        Text13.Text = ""
        MsgBox "This code does not exist in Master", vbInformation, "Code Not Found"
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
    'Text11.SetFocus
    Exit Sub
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


Private Sub Text11_GotFocus()
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


End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(Text11.Text) <> "" Then
'            If pressf1 = True Then
'                Text12.Text = Format(tmp(List3.ListIndex), "dd/mm/yyyy")
'                Text11.Text = List3.List(List3.ListIndex)

'checkallot
'If dallot > 0 Then
                DO_data
                pressf1 = False
'Else
'    MsgBox "Allotment for the day exceeded or no allotment"
'End If

            Else
                DOSHOW
            End If
    End If

End Sub

Private Sub Text11_KeyUp(KeyCode As Integer, Shift As Integer)
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
    'List3.SetFocus
End If

End Sub

Private Sub Text12_Change()
Text14.Text = Left(Text12.Text, 4) + "XXXXXXXXXXXXXXXXXX" + Right(Text12.Text, 4)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text13.Locked = False
Else
Text13.Locked = True
End If

If KeyAscii = 13 Then
'Text11.SetFocus
End If
End Sub

Private Sub Text13_LostFocus()
If Trim(Text13.Text) = "" Then
    MsgBox "Please fill RLW"
    'Text13.SetFocus
ElseIf IsNumeric(Trim(Text13.Text)) = False Then
    Text13.Text = ""
    MsgBox "Please fill numeric value for RLW"
    'Text13.SetFocus
End If
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text2.Locked = False
Else
Text2.Locked = True
End If

If KeyAscii = 13 Then
'Text4.SetFocus
End If
End Sub


Private Sub Text24_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text24.Locked = False
Else
Text24.Locked = True
End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text27.Locked = False
Else
Text27.Locked = True
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 13 Then
Text3.Locked = False
Else
Text3.Locked = True
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Then
Text4.Locked = False
Else
Text4.Locked = True
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
'Text6.SetFocus
Else
Call KeyPress1(KeyAscii)

End If
End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Text17.SetFocus
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
'Text17.SetFocus
Exit Sub
End If
End Sub


Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
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


Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(Text9.Text) > 0 Then
'cmdsave.SetFocus
Else
If Command1.Enabled = True Then Command1.SetFocus
End If
End If

End Sub


Private Sub Text17_GotFocus()
collhlp
'List1.Visible = False
Text17.SelStart = 0
Text17.SelLength = Len(Text17.Text)
For i = 0 To List3.ListCount - 1
      If Trim(Text17.Text) = Left(List2.List(i), Len(Trim(Text17.Text))) Then
                    List3.ListIndex = i
                    Exit Sub
        End If
    Next i


End Sub

Private Sub itmhlp()
'Set rs = New ADODB.Recordset
'rs.Open "Select * from mater1 order by m_name", co, adOpenKeyset, adLockOptimistic
'If rs.RecordCount > 0 Then
'rs.MoveFirst
'End If
'g = 0
'ReDim tmcode(rs.RecordCount)
'While Not rs.EOF
'List2.AddItem rs.Fields("m_name").Value
'tmcode(g) = rs.Fields("m_code").Value
'g = g + 1
'rs.MoveNext
'Wend
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

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text17.Text) <> "" Then
        If pressf1 = True Then
            Text16.Text = List2.List(List2.ListIndex)
            Text17.Text = tmcode1(List2.ListIndex)
            pressf1 = False
        Else
            colldatashow
        End If
        If Command1.Enabled = True Then Command1.SetFocus
    End If
End If

End Sub



Private Sub Timer1_Timer()
gettag1
End Sub
