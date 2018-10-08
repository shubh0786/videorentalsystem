VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmPenaltySettings 
   Caption         =   "DISC Penalty Settings"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   5895
      _Version        =   851972
      _ExtentX        =   10398
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton bttnClose 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmPenaltySettings.frx":0000
      End
      Begin XtremeSuiteControls.PushButton bttnEdit 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmPenaltySettings.frx":039A
      End
      Begin XtremeSuiteControls.PushButton bttnSave 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmPenaltySettings.frx":0734
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _Version        =   851972
      _ExtentX        =   10398
      _ExtentY        =   9975
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "GroupBox8"
      Item(0).Control(1)=   "GroupBox3"
      Item(0).Control(2)=   "GroupBox4"
      Item(0).Control(3)=   "GroupBox7"
      Item(1).Caption =   "Item"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "GroupBox9"
      Item(1).Control(1)=   "GroupBox2"
      Item(1).Control(2)=   "GroupBox5"
      Item(1).Control(3)=   "GroupBox6"
      Begin XtremeSuiteControls.GroupBox GroupBox8 
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   3413
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.GroupBox GroupBox10 
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   4935
            _Version        =   851972
            _ExtentX        =   8705
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "DVD"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.ComboBox cboDueImpose 
               Height          =   360
               Left            =   2280
               TabIndex        =   7
               Top             =   840
               Width           =   2535
               _Version        =   851972
               _ExtentX        =   4471
               _ExtentY        =   635
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit dvdDue 
               Height          =   375
               Left            =   2280
               TabIndex        =   8
               Top             =   360
               Width           =   2535
               _Version        =   851972
               _ExtentX        =   4471
               _ExtentY        =   661
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
            End
            Begin XtremeSuiteControls.Label Label11 
               Height          =   375
               Left            =   480
               TabIndex        =   10
               Top             =   360
               Width           =   1695
               _Version        =   851972
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Penalty Amount:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label12 
               Height          =   375
               Left            =   840
               TabIndex        =   9
               Top             =   840
               Width           =   1455
               _Version        =   851972
               _ExtentX        =   2566
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Effective on:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.Label Label13 
            Height          =   615
            Left            =   1800
            TabIndex        =   11
            Top             =   0
            Width           =   1455
            _Version        =   851972
            _ExtentX        =   2566
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Due Rate"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Damage"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit dvdDamage 
            Height          =   375
            Left            =   2400
            TabIndex        =   13
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   375
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Lose"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit dvdLose 
            Height          =   375
            Left            =   2400
            TabIndex        =   16
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   375
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   4320
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Tampering"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit dvdTampering 
            Height          =   375
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   375
            Left            =   720
            TabIndex        =   20
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox9 
         Height          =   1935
         Left            =   -69760
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   3413
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.GroupBox GroupBox11 
            Height          =   1335
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   4935
            _Version        =   851972
            _ExtentX        =   8705
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "VCD"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.ComboBox cboVcdEffective 
               Height          =   315
               Left            =   2280
               TabIndex        =   23
               Top             =   840
               Width           =   2535
               _Version        =   851972
               _ExtentX        =   4471
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit vcdDue 
               Height          =   375
               Left            =   2280
               TabIndex        =   24
               Top             =   360
               Width           =   2535
               _Version        =   851972
               _ExtentX        =   4471
               _ExtentY        =   661
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   375
               Left            =   960
               TabIndex        =   26
               Top             =   840
               Width           =   1335
               _Version        =   851972
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Effective on:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   375
               Left            =   600
               TabIndex        =   25
               Top             =   360
               Width           =   1575
               _Version        =   851972
               _ExtentX        =   2778
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Penalty Amount:"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.Label Label16 
            Height          =   615
            Left            =   1800
            TabIndex        =   27
            Top             =   0
            Width           =   1455
            _Version        =   851972
            _ExtentX        =   2566
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Due Rate"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   735
         Left            =   -69760
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Damage"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vcdDamage 
            Height          =   375
            Left            =   2400
            TabIndex        =   29
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   375
            Left            =   720
            TabIndex        =   30
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   735
         Left            =   -69760
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Lose"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vcdLose 
            Height          =   375
            Left            =   2400
            TabIndex        =   32
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   375
            Left            =   720
            TabIndex        =   33
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   735
         Left            =   -69760
         TabIndex        =   34
         Top             =   4320
         Visible         =   0   'False
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Tampering"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vcdTampering 
            Height          =   375
            Left            =   2400
            TabIndex        =   35
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   375
            Left            =   720
            TabIndex        =   36
            Top             =   240
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Penalty Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmPenaltySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rType As String


Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
Call unlockF
bttnSave.Enabled = True
Me.bttnEdit.Enabled = False
End Sub

Private Sub bttnSave_Click()
rType = ""
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_PenaltyRate", cn, 1, 2
    
    If .EOF = True Then
        .AddNew
        !dvdDue = Val(dvdDue.Text)
        !dvdLose = Val(dvdLose.Text)
        !DVDEffective = Me.cboDueImpose.Text
        !dvdDamage = Val(Me.dvdDamage.Text)
        !dvdTampering = Val(Me.dvdTampering.Text)
        !vcdDue = Val(Me.vcdDue.Text)
        !VCDEffective = Me.cboVcdEffective.Text
        !vcdDamage = Val(Me.vcdDamage.Text)
        !vcdLose = Val(Me.vcdLose.Text)
        !vcdTampering = Val(Me.vcdTampering.Text)
        
        .Update
        
        rType = "Add Penalty"
      
    Else
        !dvdDue = Val(dvdDue.Text)
        !dvdLose = Val(dvdLose.Text)
        !DVDEffective = Me.cboDueImpose.Text
        !dvdDamage = Val(Me.dvdDamage.Text)
        !dvdTampering = Val(Me.dvdTampering.Text)
        !vcdDue = Val(Me.vcdDue.Text)
        !VCDEffective = Me.cboVcdEffective.Text
        !vcdDamage = Val(Me.vcdDamage.Text)
        !vcdLose = Val(Me.vcdLose.Text)
        !vcdTampering = Val(Me.vcdTampering.Text)
        
        .Update
        
        rType = "Edit Penalty"
     
    End If
    MsgBox "Penalty Settings Saved.", vbInformation + vbOKOnly
    .Close
End With
    Me.bttnSave.Enabled = False
    Me.bttnEdit.Enabled = True
    Me.bttnEdit.Caption = "&Edit"
End Sub

'Sub transReport()
'Set rs = New ADODB.Recordset
'ctr = 1
'With rs
'    .Open "Select * from tbl_Transaction", cn, 1, 2
'
'    If .EOF = True Then
'        ctr = 1
'    Else
'        Do Until .EOF
'            ctr = ctr + 1
'            .MoveNext
'        Loop
'    End If
'
'    .AddNew
'    !TransID = ctr
'    !TransType = rType
'    !UserID = user
'    !TDate = Date
'    !TTime = Time
'    .Update
'
'    .Close
'End With
'End Sub


Private Sub dvdDamage_Change()
    If Me.dvdDamage = "" Or Me.dvdDamage = "0" Then Me.dvdDamage.Text = "0.00"
End Sub

Private Sub dvdDue_Change()
    If Me.dvdDue = "" Or Me.dvdDue = "0" Then Me.dvdDue.Text = "0.00"
End Sub

Private Sub dvdLose_Change()
    If Me.dvdLose = "" Or Me.dvdLose = "0" Then Me.dvdLose.Text = "0.00"
End Sub

Private Sub dvdTampering_Change()
    If Me.dvdTampering = "" Or Me.dvdTampering = "0" Then Me.dvdTampering.Text = "0.00"
End Sub

Private Sub vcdDamage_Change()
    If Me.vcdDamage = "" Or Me.vcdDamage = "0" Then Me.vcdDamage.Text = "0.00"
End Sub

Private Sub vcdDue_Change()
    If Me.vcdDue = "" Or Me.vcdDue = "0" Then Me.vcdDue.Text = "0.00"
End Sub

Private Sub vcdLose_Change()
    If Me.vcdLose = "" Or Me.vcdLose = "0" Then Me.vcdLose.Text = "0.00"
End Sub

Private Sub vcdTampering_Change()
    If Me.vcdTampering = "" Or Me.vcdTampering = "0" Then Me.vcdTampering.Text = "0.00"
End Sub

Private Sub Form_Load()
    Call clearF

    Me.cboDueImpose.AddItem "Day After the Due Date"
    Me.cboDueImpose.AddItem "After 24 hours of Rentals"
    
    Me.cboVcdEffective.AddItem "Day After the Due Date"
    Me.cboVcdEffective.AddItem "After 24 hours of Rentals"
    
    Call loadRecords
    Call lockF
    Me.bttnSave.Enabled = False
    Me.bttnEdit.Enabled = True

    Me.TabControl1(0).Caption = "DVD Penalty Rate"
    Me.TabControl1(1).Caption = "VCD Penalty Rate"
End Sub

Sub lockF()
With Me
    .dvdDamage.Locked = True
    .dvdDue.Locked = True
    .dvdLose.Locked = True
    .dvdTampering.Locked = True
    .cboDueImpose.Locked = True
    .vcdDamage.Locked = True
    .vcdDue.Locked = True
    .vcdLose.Locked = True
    .vcdTampering.Locked = True
    .cboVcdEffective.Locked = True
End With
End Sub

Sub unlockF()
With Me
    .dvdDamage.Locked = False
    .dvdDue.Locked = False
    .dvdLose.Locked = False
    .dvdTampering.Locked = False
    .cboDueImpose.Locked = False
    .vcdDamage.Locked = False
    .vcdDue.Locked = False
    .vcdLose.Locked = False
    .vcdTampering.Locked = False
    .cboVcdEffective.Locked = False
End With
End Sub

Sub loadRecords()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_PenaltyRate", cn, 1, 2
    If .EOF = True Then
        Me.bttnEdit.Caption = "&New"
    Else
        Me.bttnEdit.Caption = "&Edit"
    End If
    
    Do While .EOF = False
        dvdDue.Text = !dvdDue
        dvdLose.Text = !dvdLose
        Me.cboDueImpose.Text = !DVDEffective
        Me.dvdDamage.Text = !dvdDamage
        Me.dvdTampering.Text = !dvdTampering
        Me.vcdDue.Text = !vcdDue
        Me.cboVcdEffective.Text = !VCDEffective
        Me.vcdDamage.Text = !vcdDamage
        Me.vcdLose.Text = !vcdLose
        Me.vcdTampering.Text = !vcdTampering
        .MoveNext
    Loop
    .Close
End With
End Sub

Sub clearF()
With Me
    .dvdDamage = ""
    .dvdDue = ""
    .dvdLose = ""
    .dvdTampering = ""
    .cboDueImpose = ""
    .vcdDamage = ""
    .vcdDue = ""
    .vcdLose = ""
    .vcdTampering = ""
    .cboVcdEffective = ""
End With
End Sub

