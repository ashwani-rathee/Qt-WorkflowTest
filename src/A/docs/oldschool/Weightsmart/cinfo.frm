VERSION 5.00
Begin VB.Form cinfo 
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
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
      Height          =   5235
      Left            =   9360
      TabIndex        =   14
      Top             =   1620
      Width           =   3255
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
      Left            =   12360
      TabIndex        =   9
      ToolTipText     =   "Unload Form"
      Top             =   600
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   11895
      Begin VB.TextBox Text60 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   69
         Text            =   " Machine7"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Text46 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   68
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Text47 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   67
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox Text48 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   66
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Text49 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   65
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox Text39 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   30
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox Text38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   29
         Top             =   4740
         Width           =   1815
      End
      Begin VB.TextBox Text37 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   28
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox Text36 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   27
         Top             =   4740
         Width           =   1815
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   49
         Text            =   " Controller (IN)"
         Top             =   4740
         Width           =   1815
      End
      Begin VB.TextBox Text34 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   26
         Top             =   4380
         Width           =   1335
      End
      Begin VB.TextBox Text33 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   25
         Top             =   4380
         Width           =   1815
      End
      Begin VB.TextBox Text32 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   24
         Top             =   4380
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   23
         Top             =   4380
         Width           =   1815
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   48
         Text            =   " RFID (IN)"
         Top             =   4380
         Width           =   1815
      End
      Begin VB.TextBox Text29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   34
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   33
         Top             =   3420
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   32
         Top             =   3420
         Width           =   1575
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   31
         Top             =   3420
         Width           =   2055
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   47
         Text            =   " DO Allotment"
         Top             =   3420
         Width           =   1815
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   22
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   21
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   20
         Top             =   3060
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   19
         Top             =   3060
         Width           =   2055
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   45
         Text            =   " Camera"
         Top             =   3060
         Width           =   1815
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   38
         Top             =   2700
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   37
         Top             =   2700
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   36
         Top             =   2700
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   35
         Top             =   2700
         Width           =   2055
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
         Height          =   405
         Left            =   6840
         TabIndex        =   18
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   17
         Top             =   2340
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   16
         Top             =   2340
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   15
         Top             =   2340
         Width           =   2055
      End
      Begin VB.TextBox Text5 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text4 
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
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
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
         TabIndex        =   8
         Top             =   5880
         Width           =   3255
      End
      Begin VB.TextBox Text3 
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
         Top             =   1440
         Width           =   4575
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text1 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   660
         Width           =   4575
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   6840
         TabIndex        =   42
         Text            =   " Password"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   5520
         TabIndex        =   41
         Text            =   " Username"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   3960
         TabIndex        =   40
         Text            =   " Path"
         Top             =   1980
         Width           =   1575
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   1920
         TabIndex        =   39
         Text            =   " IP Address"
         Top             =   1980
         Width           =   2055
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   46
         Text            =   " Headquarters"
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   44
         Text            =   " Area"
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   120
         TabIndex        =   43
         Top             =   1980
         Width           =   1815
      End
      Begin VB.TextBox Text41 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   6840
         TabIndex        =   50
         Text            =   " Local Host"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text42 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   5040
         TabIndex        =   51
         Text            =   " Local IP"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text43 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   3720
         TabIndex        =   52
         Text            =   " Host"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text44 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   1920
         TabIndex        =   53
         Text            =   " Remote IP"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text45 
         BackColor       =   &H00E0E0E0&
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
         Height          =   405
         Left            =   120
         TabIndex        =   54
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox Text59 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   55
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox Text58 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   56
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text57 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   57
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox Text56 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   58
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text50 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   59
         Text            =   " Controller (OUT)"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text54 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6840
         TabIndex        =   60
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox Text53 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   61
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox Text52 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   62
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox Text51 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   63
         Top             =   5160
         Width           =   1815
      End
      Begin VB.TextBox Text55 
         BackColor       =   &H80000004&
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
         Height          =   405
         Left            =   120
         TabIndex        =   64
         Text            =   " RFID (OUT)"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Area Code"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "WBridge Code"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Area Details"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Company"
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
         TabIndex        =   3
         Top             =   660
         Width           =   1575
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Company Information"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000A0&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   11895
   End
End
Attribute VB_Name = "cinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idd As Integer

Private Sub Command1_Click()
Dim ctr As Integer

Set rs1 = New ADODB.Recordset

a = MsgBox("Want to Update Records", vbOKCancel, "Update Windows")
If a = 1 Then

rs1.Open "select * from party ", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Else
rs1.AddNew
End If

rs1.Fields("AREACODE").Value = Text5.Text & ""
rs1.Fields("WBCODE").Value = Text4.Text & ""
rs1.Fields("pname").Value = Text1.Text & ""
rs1.Fields("pinf").Value = Text2.Text & ""
rs1.Fields("Padd").Value = Text3.Text & ""
rs1.Update

Set rs2 = New ADODB.Recordset
rs2.Open "Select * from paths order by id", co, adOpenKeyset, adLockOptimistic
rs2.MoveLast
ctr = Val(rs2.Fields("id")) + 1

While rs2.RecordCount < 9
rs2.AddNew
rs2.Fields("id").Value = ctr
ctr = ctr + 1
rs2.Update
Wend

If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    rs2.Fields("ipaddress").Value = Text6.Text
    rs2.Fields("path").Value = Text7.Text
    rs2.Fields("username").Value = Text8.Text
    rs2.Fields("password").Value = Text9.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text10.Text
    rs2.Fields("path").Value = Text11.Text
    rs2.Fields("username").Value = Text12.Text
    rs2.Fields("password").Value = Text13.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text14.Text
    rs2.Fields("path").Value = Text15.Text
    rs2.Fields("username").Value = Text16.Text
    rs2.Fields("password").Value = Text17.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text26.Text
    rs2.Fields("path").Value = Text27.Text
    rs2.Fields("username").Value = Text28.Text
    rs2.Fields("password").Value = Text29.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text31.Text
    rs2.Fields("path").Value = Text32.Text
    rs2.Fields("username").Value = Text33.Text
    rs2.Fields("password").Value = Text34.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text36.Text
    rs2.Fields("path").Value = Text37.Text
    rs2.Fields("username").Value = Text38.Text
    rs2.Fields("password").Value = Text39.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text51.Text
    rs2.Fields("path").Value = Text52.Text
    rs2.Fields("username").Value = Text53.Text
    rs2.Fields("password").Value = Text54.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text56.Text
    rs2.Fields("path").Value = Text57.Text
    rs2.Fields("username").Value = Text58.Text
    rs2.Fields("password").Value = Text59.Text
    rs2.MoveNext
    rs2.Fields("ipaddress").Value = Text46.Text
    rs2.Fields("path").Value = Text47.Text
    rs2.Fields("username").Value = Text48.Text
    rs2.Fields("password").Value = Text49.Text
End If
rs2.Update
rs2.Close
MsgBox "Updated"
End If
End Sub

Private Sub Command1_GotFocus()
If Trim(Text1.Text) = "" Then
Text1.SetFocus
Exit Sub
End If


End Sub

Private Sub Command4_Click()
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
Me.Picture = main.Picture
Call Conn
compinfor
wblist
Set rs = New ADODB.Recordset
rs.Open "Select * from paths order by id", co, adOpenKeyset, adLockOptimistic
While rs.RecordCount < 4
rs.MoveLast
idd = rs.Fields(0).Value
rs.AddNew
rs.Fields(0) = idd + 1
rs.Fields(1) = ""
rs.Fields(2) = ""
rs.Fields(3) = ""
rs.Fields(4) = ""
rs.Update
Wend
If rs.RecordCount > 0 Then
    rs.MoveFirst
    Text6.Text = rs.Fields("ipaddress").Value
    Text7.Text = rs.Fields("path").Value
    Text8.Text = rs.Fields("username").Value
    Text9.Text = rs.Fields("password").Value
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text10.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text11.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text12.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text13.Text = rs.Fields("password").Value
    End If
    rs.MoveNext
    Text14.Text = rs.Fields("ipaddress").Value
    Text15.Text = rs.Fields("path").Value
    Text16.Text = rs.Fields("username").Value
    Text17.Text = rs.Fields("password").Value
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text26.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text27.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text28.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text29.Text = rs.Fields("password").Value
    End If
    
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text31.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text32.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text33.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text34.Text = rs.Fields("password").Value
    End If
    
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text36.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text37.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text38.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text39.Text = rs.Fields("password").Value
    End If
    
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text51.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text52.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text53.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text54.Text = rs.Fields("password").Value
    End If
    
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text56.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text57.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text58.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text59.Text = rs.Fields("password").Value
    End If
    
    rs.MoveNext
    If Trim(rs.Fields("ipaddress").Value) <> "" Then
        Text46.Text = rs.Fields("ipaddress").Value
    End If
    If Trim(rs.Fields("path").Value) <> "" Then
        Text47.Text = rs.Fields("path").Value
    End If
    If Trim(rs.Fields("username").Value) <> "" Then
        Text48.Text = rs.Fields("username").Value
    End If
    If Trim(rs.Fields("password").Value) <> "" Then
        Text49.Text = rs.Fields("password").Value
    End If
End If
rs.Close
End Sub


Private Sub wblist()
If List1.ListCount > 0 Then
List1.Clear
List1.Refresh
End If

Set rs = New ADODB.Recordset
rs.Open "Select * from wbcode order by wbcode", co, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
rs.MoveFirst
g = 0

End If
g = 0
ReDim tmpcode(rs.RecordCount)
ReDim tmpcode1(rs.RecordCount)
While Not rs.EOF
List1.AddItem padl(Trim(rs.Fields("wbcode").Value), 3) & " | " & padl(Trim(rs.Fields("wbname").Value), 50)
tmpcode(g) = rs.Fields("wbcode").Value
tmpcode1(g) = rs.Fields("wbname").Value
g = g + 1
rs.MoveNext
Wend
End Sub

Private Sub compinfor()
Set rs1 = New ADODB.Recordset
rs1.Open "Select * from party ", co, adOpenKeyset, adLockOptimistic
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Text5.Text = rs1.Fields("AREACODE").Value & ""
Text4.Text = rs1.Fields("WBCODE").Value & ""
Text1.Text = rs1.Fields("pname").Value & ""
Text2.Text = rs1.Fields("pinf").Value & ""
Text3.Text = rs1.Fields("Padd").Value & ""
End If

End Sub


Private Sub List1_DblClick()
Dim awcode As String
awcode = padl(List1.List(List1.ListIndex), 3)
Text5.Text = Left(awcode, 2)
Text4.Text = Right(awcode, 1)
awcode = Right(List1.List(List1.ListIndex), 50)
Text2.Text = awcode
Text3.Text = "Area"
End Sub
