VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmRent 
   Caption         =   "Rent Disc"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _Version        =   851972
      _ExtentX        =   22675
      _ExtentY        =   14843
      _StockProps     =   79
      Caption         =   "Renting Process"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Begin XtremeSuiteControls.ListView ListView1 
         Height          =   570
         Left            =   6240
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   6375
         _Version        =   851972
         _ExtentX        =   11245
         _ExtentY        =   1005
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListBox lstMember 
         Height          =   135
         Left            =   2040
         TabIndex        =   40
         Top             =   5160
         Visible         =   0   'False
         Width           =   3975
         _Version        =   851972
         _ExtentX        =   7011
         _ExtentY        =   238
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
      End
      Begin XtremeSuiteControls.ListBox lstTitle 
         Height          =   150
         Left            =   2040
         TabIndex        =   39
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   265
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
      End
      Begin XtremeSuiteControls.ListBox lstInDi 
         Height          =   135
         Left            =   2040
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   238
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
      End
      Begin XtremeSuiteControls.FlatEdit txtIndi 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   12582912
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDVDNo 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTitle 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGenre 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2400
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtActor 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2880
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCDtype 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   3360
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRentFee 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   3840
         Width           =   3855
         _Version        =   851972
         _ExtentX        =   6800
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMemberID 
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   4800
         Width           =   3975
         _Version        =   851972
         _ExtentX        =   7011
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtName 
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   5280
         Width           =   3975
         _Version        =   851972
         _ExtentX        =   7011
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   2535
         Left            =   6960
         TabIndex        =   20
         Top             =   3000
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   4471
         _StockProps     =   79
         Appearance      =   6
         Begin XtremeSuiteControls.FlatEdit txtTotalAmount 
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            Top             =   840
            Width           =   2655
            _Version        =   851972
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtItem 
            Height          =   375
            Left            =   2400
            TabIndex        =   22
            Top             =   360
            Width           =   2655
            _Version        =   851972
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPayment 
            Height          =   375
            Left            =   2640
            TabIndex        =   26
            Top             =   1440
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtChange 
            Height          =   375
            Left            =   2640
            TabIndex        =   28
            Top             =   1800
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   495
            Left            =   1440
            TabIndex        =   27
            Top             =   1800
            Width           =   975
            _Version        =   851972
            _ExtentX        =   1720
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Change:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   495
            Left            =   1320
            TabIndex        =   25
            Top             =   1440
            Width           =   1215
            _Version        =   851972
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Payment:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label20 
            Height          =   495
            Left            =   720
            TabIndex        =   24
            Top             =   840
            Width           =   1695
            _Version        =   851972
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Total Amount:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label21 
            Height          =   495
            Left            =   720
            TabIndex        =   23
            Top             =   360
            Width           =   1695
            _Version        =   851972
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Number of Item:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton bttnAddItem 
         Height          =   495
         Left            =   1320
         TabIndex        =   29
         Top             =   6600
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Add Item"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmRent.frx":0000
      End
      Begin XtremeSuiteControls.PushButton bttnRemove 
         Height          =   495
         Left            =   4440
         TabIndex        =   30
         Top             =   6600
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Remove Item"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmRent.frx":039A
      End
      Begin XtremeSuiteControls.PushButton bttnCheckOut 
         Height          =   495
         Left            =   7200
         TabIndex        =   31
         Top             =   6600
         Width           =   1935
         _Version        =   851972
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "OK"
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
         Picture         =   "frmRent.frx":0734
      End
      Begin XtremeSuiteControls.PushButton bttnCancel 
         Height          =   495
         Left            =   9720
         TabIndex        =   32
         Top             =   6600
         Width           =   2175
         _Version        =   851972
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cancel"
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
         Picture         =   "frmRent.frx":0ACE
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   7560
         Width           =   6495
         _Version        =   851972
         _ExtentX        =   11456
         _ExtentY        =   1296
         _StockProps     =   79
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   120
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Date"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label lblDate 
            Height          =   495
            Left            =   960
            TabIndex        =   36
            Top             =   120
            Width           =   1935
            _Version        =   851972
            _ExtentX        =   3413
            _ExtentY        =   873
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   495
            Left            =   2880
            TabIndex        =   35
            Top             =   120
            Width           =   615
            _Version        =   851972
            _ExtentX        =   1085
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Time"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label lblTime 
            Height          =   495
            Left            =   3600
            TabIndex        =   34
            Top             =   120
            Width           =   1935
            _Version        =   851972
            _ExtentX        =   3413
            _ExtentY        =   873
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   1575
         Left            =   0
         TabIndex        =   41
         Top             =   6000
         Width           =   12735
         _Version        =   851972
         _ExtentX        =   22463
         _ExtentY        =   2778
         _StockProps     =   68
         Appearance      =   11
         Color           =   128
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   5280
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Member Name:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label18 
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   4800
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Member ID:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   3840
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Rental Fee:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   3360
         Width           =   975
         _Version        =   851972
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "CD Type:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label16 
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   2880
         Width           =   1455
         _Version        =   851972
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Actor/Actress:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label15 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2400
         Width           =   735
         _Version        =   851972
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Genre:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label14 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1920
         Width           =   615
         _Version        =   851972
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Title:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label13 
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "CD/DVD ID:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Individual No.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ID As Boolean
Dim vType As String
Dim goodFor As Integer

Private Sub bttnAddItem_Click()
    Dim x, loanID As String
    Dim C As Integer
    Me.ListView1.Visible = True
    ctr = 1
    Me.ListView1.Visible = True

    If Me.txtItem.Text = "" Then C = 0
    If Me.txtItem.Text <> "" Then C = CInt(Me.txtItem.Text)
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Loan", cn, 1, 2
    x = Format(Date, "mm-dd-yyyy")
    If .EOF = True Then
        ctr = 1
        loanID = x & "-" & ctr
    Else
        Do Until .EOF
            ctr = ctr + 1
            .MoveNext
        Loop
        loanID = x & "-" & ctr
    End If
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_CDList where CD_ID like '" & Me.txtDVDNo.Text & "'", cn, 1, 2
            goodFor = !goodFor
        .Close
    End With
    
    Set lst = Me.ListView1.ListItems.Add(, , loanID)
        With lst
            .ListSubItems.Add , , Me.txtIndi.Text
            .ListSubItems.Add , , Me.txtTitle.Text
            .ListSubItems.Add , , Now
            .ListSubItems.Add , , DateAdd("d", goodFor, Date)
            .ListSubItems.Add , , Me.txtMemberID.Text
            .ListSubItems.Add , , Me.txtCDtype.Text
            .ListSubItems.Add , , Me.txtRentFee.Text
            
            Me.txtTotalAmount.Text = CInt(Me.txtTotalAmount.Text) + CInt(Me.txtRentFee.Text)
        End With
        Me.ListView1.Enabled = True
        If Me.ListView1.Height >= 2085 Then
        Else
            Me.ListView1.Height = Me.ListView1.Height + 300
        End If
            
            C = C + 1
            Me.txtItem.Text = C
            
            With Me
                .txtActor.Text = ""
                .txtDVDNo.Text = ""
                .txtTitle.Text = ""
                .txtIndi.Text = ""
                .txtGenre.Text = ""
                .txtCDtype.Text = ""
                .txtRentFee.Text = ""
            End With
    
    .Close
End With
Me.txtMemberID.Locked = True
End Sub

Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnCheckOut_Click()
    Me.txtMemberID.Locked = False
    Call getIncomeID

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Loan", cn, 1, 2
    For I = 1 To CInt(Me.txtItem.Text)
        .AddNew
        Set lst = Me.ListView1.ListItems.Item(I)
            !loanID = Me.ListView1.ListItems.Item(I)
            !Ino = lst.ListSubItems(1)
            !rentDate = lst.ListSubItems(3)
            !ReturnDate = lst.ListSubItems(4)
            !MemberID = lst.ListSubItems(5)
            !Status = "Rent"
            !rentfee = True
            !Duefee = False
            !damageFee = False
            !temperedFee = False
            !LoseFee = False
            .Update
            
            Set rs2 = New ADODB.Recordset
            With rs2
                .Open "Select * from tbl_IndividualCDList where Ino like '" & lst.ListSubItems(1) & "'", cn, 1, 2
                    !Status = "RENTED"
                    .Update
                .Close
            End With
            
           
            wDone = "Borrow of CD" & "/" & "Loan ID" & " - " & Me.ListView1.ListItems.Item(I) & "/" & "CD Number" & " - " & lst.ListSubItems(1)
            Call auditT

    Next I
    
          .Close

        Set rs = New ADODB.Recordset

            .Open "Select * from tbl_Income", cn, 1, 2
            .AddNew
            !incomeID = incomeID
            !amount = Me.txtTotalAmount.Text
            !iDate = Now
            !iType = "RENTALS"
            !UserID = user
            .Update
            .Close
        
            MsgBox "Record Save.", vbInformation + vbOKOnly
            Me.ListView1.ListItems.clear
            Me.txtTotalAmount.Text = ""
            Me.txtItem.Text = ""
            Me.txtPayment.Text = ""
            Me.txtChange.Text = ""
            Me.txtMemberID.Text = ""
            Me.txtName.Text = ""
    
End With


End Sub

Private Sub Form_Load()
    Me.lblDate.Caption = Date
    Me.lblTime.Caption = Time
    
    Me.Top = 1700
    Me.Left = 3250
    
    Me.ListView1.ColumnHeaders.clear
    Me.ListView1.ListItems.clear
    Me.ListView1.ColumnHeaders.Add , , "Loan ID", 1500
    Me.ListView1.ColumnHeaders.Add , , "CD Number", 1500
    Me.ListView1.ColumnHeaders.Add , , "Title", 2000
    Me.ListView1.ColumnHeaders.Add , , "Rent Date", 1500
    Me.ListView1.ColumnHeaders.Add , , "Return Date", 1500
    Me.ListView1.ColumnHeaders.Add , , "Member ID", 1500
    Me.ListView1.ColumnHeaders.Add , , "CD Type", 1500
    Me.ListView1.ColumnHeaders.Add , , "Rental Fee", 1500
    
    
    Me.txtTotalAmount.Text = "0"
    Me.txtItem.Text = "0"
    Me.bttnCheckOut.Enabled = False
End Sub



Private Sub ListView1_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Me.bttnRemove.Enabled = True
End Sub

Private Sub lstInDi_Click()
x = Me.lstInDi.ListIndex
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from QViewAVAILABLE where INo like '" & Me.lstInDi.List(x) & "'", cn, 1, 2
        Me.txtDVDNo.Text = !CD_ID
        Me.txtIndi.Text = !Ino
        Me.txtActor.Text = !actor
        Me.txtGenre.Text = !Ggenre
        Me.txtTitle.Text = !Title
    .Close
    Me.lstInDi.Visible = False
    Me.txtDVDNo.Enabled = False
End With

    Set rs = New ADODB.Recordset
    With rs
        .Open "Select CDTYpe from tbl_CDList where CD_ID like '" & Me.txtDVDNo.Text & "'", cn, 1, 2
            vType = !CDType
            Me.txtCDtype.Text = !CDType
        .Close
    End With
    
        Set rs = New ADODB.Recordset
        With rs
            If vType = "VCD" Then
                .Open "Select VCDAmount from tbl_RentalRate", cn, 1, 2
                Me.txtRentFee.Text = !VCDAmount
            Else
                .Open "Select DVDAmount from tbl_RentalRate", cn, 1, 2
                Me.txtRentFee.Text = !DVDAmount
            End If
            .Close
        End With
End Sub

Private Sub lstMember_Click()
x = Me.lstMember.ListIndex
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Members where MemberID like '" & Me.lstMember.List(x) & "'", cn, 1, 2
    If .EOF = True Then
        MsgBox "No Record Found.", vbInformation + vbOKOnly
    Else
        Me.txtName.Text = !Lname & ", " & !Fname & " " & !Mname
        Me.lstMember.Visible = False
    End If
    .Close
End With
End Sub

Private Sub lstTitle_Click()
Dim a() As String
Dim sltCtr As Integer

x = Me.lstTitle.ListIndex
    a = Split(Me.lstTitle.List(x), " - ")
    
    For sltCtr = 0 To UBound(a)
    
    Next sltCtr

Set rs = New ADODB.Recordset
With rs
    .Open "select * from QViewAVAILABLE where Ino like '" & a(1) & "'", cn, 1, 2
    
    Me.txtIndi.Text = a(1)
    Me.txtActor.Text = !actor
    Me.txtDVDNo.Text = !CD_ID
    Me.txtTitle.Text = !Title
    Me.txtGenre.Text = !Ggenre
    .Close
    Me.lstTitle.Visible = False
    Me.lstInDi.Visible = False
    '.Close
End With
    
    
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select CDTYpe from tbl_CDList where CD_ID like '" & Me.txtDVDNo.Text & "'", cn, 1, 2
            vType = !CDType
            Me.txtCDtype.Text = !CDType
        .Close
    End With
    
        Set rs = New ADODB.Recordset
        With rs
            If vType = "VCD" Then
                .Open "Select VCDAmount from tbl_RentalRate", cn, 1, 2
                Me.txtRentFee.Text = !VCDAmount
            Else
                .Open "Select DVDAmount from tbl_RentalRate", cn, 1, 2
                Me.txtRentFee.Text = !DVDAmount
            End If
            .Close
        End With
End Sub



Private Sub txtChange_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtIndi_Change()
If ID = True Then
    If Me.txtIndi.Text = "" Then
        Me.lstInDi.Visible = False
    Else
        x = 0
        Me.lstInDi.clear
        Me.lstInDi.Visible = True
        Set rs = New ADODB.Recordset
        With rs
            .Open "Select * from QViewAVAILABLE where INo like '" & Me.txtIndi.Text & "%'", cn, 1, 2
                If .EOF = True Then
                    MsgBox "NO RECORD of DVD.", vbInformation + vbOKOnly
                    Me.lstInDi.Visible = False
                Else
                    Do Until .EOF
                        Me.lstInDi.AddItem !Ino
                        If x >= 2500 Then
                        Else
                            x = x + 300
                        End If
                        .MoveNext
                    Loop
                    Me.lstInDi.Height = x
                End If
            .Close
        End With
    End If
End If
End Sub

Private Sub txtIndi_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        ID = True
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txtItem_Change()
    If Me.txtItem.Text = "" Then Me.txtItem.Text = "0"
    If CInt(Me.txtItem.Text) > 0 Then
        Me.txtPayment.Enabled = True
        Me.txtChange.Enabled = True
    End If
End Sub

Private Sub txtMemberID_Change()
If Me.txtMemberID.Text = "" Then
    Me.lstMember.Visible = False
Else
    x = 0
    Me.lstMember.clear
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Members where memberID like '" & Me.txtMemberID.Text & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
            If .EOF = True Then
                MsgBox "No Member Registered.", vbInformation + vbOKOnly
                Me.lstMember.Visible = False
                Me.txtMemberID.Text = ""
            Else
                Do Until .EOF
                    Me.lstMember.AddItem !MemberID
                        If x >= 2500 Then
                        Else
                            x = x + 300
                        End If
                    .MoveNext
                Loop
                Me.lstMember.Visible = True
                Me.lstMember.Height = x
            End If
        .Close
    End With
End If
End Sub

Private Sub txtMemberID_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Change()
    If Me.txtTitle.Text <> "" And Me.txtName.Text <> "" Then Me.bttnAddItem.Enabled = True
End Sub

Private Sub txtPayment_Change()
If Me.txtPayment.Text = "" Then Me.txtPayment.Text = "0"
If CInt(Me.txtPayment.Text) >= CInt(Me.txtTotalAmount.Text) Then
    Me.txtChange.Text = CInt(Me.txtPayment.Text) - CInt(Me.txtTotalAmount.Text)
    Me.bttnCheckOut.Enabled = True
Else
    Me.bttnCheckOut.Enabled = False
    Me.txtChange.Text = ""
End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = True Or KeyAscii = 8 Then

Else
    KeyAscii = 0
End If
End Sub

Private Sub txtTitle_Change()
If Me.txtTitle.Text <> "" And Me.txtName.Text <> "" Then Me.bttnAddItem.Enabled = True
If Me.txtTitle.Text = "" Or Me.txtName.Text = "" Then Me.bttnAddItem.Enabled = False
If ID = False Then
    If Me.txtTitle.Text = "" Then
        Me.lstTitle.Visible = False
    Else
        x = 0
        Me.lstTitle.clear
        Set rs = New ADODB.Recordset
        With rs
            .Open "Select * from QViewAVAILABLE where title like '" & Me.txtTitle.Text & "%'", cn, 1, 2
            If .EOF = True Then
                MsgBox "No RECORD FOUND.", vbInformation + vbOKOnly
                Me.lstTitle.Visible = False
            Else
                Do Until .EOF
                    Me.lstTitle.AddItem !Title & " - " & !Ino
                    If x >= 2500 Then
                    Else
                        x = x + 300
                    End If
                    .MoveNext
                Loop
                    Me.lstTitle.Visible = True
                    Me.lstTitle.Height = x
            End If
            .Close
        End With
    End If
End If
End Sub


Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    ID = False
End Sub

Private Sub txtTotalAmount_Change()
    If Me.txtTotalAmount.Text = "" Then Me.txtTotalAmount.Text = "0"
End Sub

Private Sub txtTotalAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

