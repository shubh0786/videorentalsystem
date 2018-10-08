VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmModifyMember 
   Caption         =   "Modify Member"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   12495
      _Version        =   851972
      _ExtentX        =   22040
      _ExtentY        =   1720
      _StockProps     =   68
      Appearance      =   8
      Begin XtremeSuiteControls.PushButton bttnCancel 
         Height          =   495
         Left            =   7200
         TabIndex        =   56
         Top             =   0
         Width           =   1935
         _Version        =   851972
         _ExtentX        =   3413
         _ExtentY        =   873
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
         Picture         =   "frmModifyMember.frx":0000
      End
      Begin XtremeSuiteControls.PushButton bttnUpdate 
         Height          =   495
         Left            =   4560
         TabIndex        =   55
         Top             =   0
         Width           =   1935
         _Version        =   851972
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Update"
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
         Picture         =   "frmModifyMember.frx":039A
      End
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12615
      _Version        =   851972
      _ExtentX        =   22251
      _ExtentY        =   12938
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
      Appearance      =   11
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label17"
      Item(0).Control(1)=   "txtMemberID"
      Item(0).Control(2)=   "Label1"
      Item(0).Control(3)=   "dtpApplication"
      Item(0).Control(4)=   "Picture1"
      Item(0).Control(5)=   "Label15"
      Item(0).Control(6)=   "txtPerAdd"
      Item(0).Control(7)=   "Label18"
      Item(0).Control(8)=   "txtPerZip"
      Item(0).Control(9)=   "Label7"
      Item(0).Control(10)=   "txtCurAdd"
      Item(0).Control(11)=   "Label8"
      Item(0).Control(12)=   "txtCurZip"
      Item(0).Control(13)=   "Label13"
      Item(0).Control(14)=   "txtCelNo"
      Item(0).Control(15)=   "Label10"
      Item(0).Control(16)=   "txtEmailAdd"
      Item(0).Control(17)=   "TabControl3"
      Item(0).Control(18)=   "bttnBrowse"
      Item(1).Caption =   "Item"
      Item(1).ControlCount=   22
      Item(1).Control(0)=   "Label25"
      Item(1).Control(1)=   "txtExLName"
      Item(1).Control(2)=   "txtExFName"
      Item(1).Control(3)=   "txtExMName"
      Item(1).Control(4)=   "Label24"
      Item(1).Control(5)=   "txtRelation"
      Item(1).Control(6)=   "Label27"
      Item(1).Control(7)=   "txtExCurAdd"
      Item(1).Control(8)=   "Label26"
      Item(1).Control(9)=   "txtExCurZip"
      Item(1).Control(10)=   "Label11"
      Item(1).Control(11)=   "txtExCelNo"
      Item(1).Control(12)=   "Label12"
      Item(1).Control(13)=   "txtExEmailAdd"
      Item(1).Control(14)=   "Label19"
      Item(1).Control(15)=   "cboID"
      Item(1).Control(16)=   "Label23"
      Item(1).Control(17)=   "txtIDNo"
      Item(1).Control(18)=   "Label20"
      Item(1).Control(19)=   "rBttn1"
      Item(1).Control(20)=   "rBttn2"
      Item(1).Control(21)=   "rBttn3"
      Begin XtremeSuiteControls.TabControl TabControl3 
         Height          =   3015
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   9015
         _Version        =   851972
         _ExtentX        =   15901
         _ExtentY        =   5318
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         Begin XtremeSuiteControls.RadioButton rBttnSingle 
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   2520
            Width           =   2055
            _Version        =   851972
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Single"
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
         End
         Begin XtremeSuiteControls.ComboBox cboGender 
            Height          =   360
            Left            =   5640
            TabIndex        =   44
            Top             =   1440
            Width           =   1575
            _Version        =   851972
            _ExtentX        =   2778
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
            Text            =   "ComboBox3"
         End
         Begin XtremeSuiteControls.FlatEdit txtLName 
            Height          =   495
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
            _ExtentY        =   873
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
            Text            =   "Last Name"
         End
         Begin XtremeSuiteControls.FlatEdit txtFName 
            Height          =   495
            Left            =   2520
            TabIndex        =   46
            Top             =   360
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
            _ExtentY        =   873
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
            Text            =   "First Name"
         End
         Begin XtremeSuiteControls.FlatEdit txtMName 
            Height          =   495
            Left            =   4920
            TabIndex        =   47
            Top             =   360
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   873
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
            Text            =   "Middle Name"
         End
         Begin XtremeSuiteControls.RadioButton rBttnMarried 
            Height          =   255
            Left            =   3120
            TabIndex        =   48
            Top             =   2520
            Width           =   2055
            _Version        =   851972
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Married"
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
         End
         Begin XtremeSuiteControls.RadioButton rBttnWidow 
            Height          =   255
            Left            =   5760
            TabIndex        =   49
            Top             =   2520
            Width           =   2055
            _Version        =   851972
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Separated"
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
         End
         Begin MSComCtl2.DTPicker dtBdate 
            Height          =   375
            Left            =   240
            TabIndex        =   57
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   46137345
            CurrentDate     =   40969
         End
         Begin XtremeSuiteControls.FlatEdit txtAge 
            Height          =   375
            Left            =   2880
            TabIndex        =   58
            Top             =   1440
            Width           =   2055
            _Version        =   851972
            _ExtentX        =   3625
            _ExtentY        =   661
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
            Alignment       =   2
            Appearance      =   6
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   45
            Width           =   2295
            _Version        =   851972
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "PRINCIPAL MEMBER"
            BackColor       =   14737632
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
            Height          =   255
            Left            =   3120
            TabIndex        =   53
            Top             =   1080
            Width           =   1815
            _Version        =   851972
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "AGE"
            BackColor       =   14737632
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Left            =   5280
            TabIndex        =   52
            Top             =   1080
            Width           =   2295
            _Version        =   851972
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "SEX"
            BackColor       =   14737632
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   1080
            Width           =   2295
            _Version        =   851972
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "DATE OF BIRTH       "
            BackColor       =   14737632
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Left            =   3360
            TabIndex        =   50
            Top             =   2040
            Width           =   1935
            _Version        =   851972
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "CIVIL STATUS          "
            BackColor       =   14737632
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   9720
         ScaleHeight     =   1965
         ScaleWidth      =   2505
         TabIndex        =   2
         Top             =   480
         Width           =   2535
         Begin VB.Image Image1 
            Height          =   1935
            Left            =   0
            Top             =   0
            Width           =   2490
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtMemberID 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton bttnBrowse 
         Height          =   495
         Left            =   10320
         TabIndex        =   4
         ToolTipText     =   "Browse Picture"
         Top             =   2520
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Browse"
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
         Picture         =   "frmModifyMember.frx":0734
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12000
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin XtremeSuiteControls.FlatEdit txtPerAdd 
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   4560
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtPerZip 
         Height          =   375
         Left            =   9960
         TabIndex        =   6
         Top             =   4560
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCurAdd 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   5040
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCurZip 
         Height          =   375
         Left            =   9960
         TabIndex        =   8
         Top             =   5040
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCelNo 
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   5640
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtEmailAdd 
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   6120
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpApplication 
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   600
         Width           =   4215
         _Version        =   851972
         _ExtentX        =   7435
         _ExtentY        =   661
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
         MinDate         =   40544
         CurrentDate     =   40544
      End
      Begin XtremeSuiteControls.FlatEdit txtExLName 
         Height          =   495
         Left            =   -68920
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851972
         _ExtentX        =   4260
         _ExtentY        =   873
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
         Text            =   "Last Name"
      End
      Begin XtremeSuiteControls.FlatEdit txtExFName 
         Height          =   495
         Left            =   -66280
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851972
         _ExtentX        =   4260
         _ExtentY        =   873
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
         Text            =   "First Name"
      End
      Begin XtremeSuiteControls.FlatEdit txtExMName 
         Height          =   495
         Left            =   -63640
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   873
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
         Text            =   "Middle Name"
      End
      Begin XtremeSuiteControls.FlatEdit txtRelation 
         Height          =   375
         Left            =   -68920
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   5055
         _Version        =   851972
         _ExtentX        =   8916
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtExCurAdd 
         Height          =   375
         Left            =   -68920
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   7215
         _Version        =   851972
         _ExtentX        =   12726
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtExCurZip 
         Height          =   375
         Left            =   -61120
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851972
         _ExtentX        =   3836
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtExCelNo 
         Height          =   375
         Left            =   -68920
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851972
         _ExtentX        =   7858
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.FlatEdit txtExEmailAdd 
         Height          =   375
         Left            =   -63400
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   4335
         _Version        =   851972
         _ExtentX        =   7646
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.ComboBox cboID 
         Height          =   360
         Left            =   -68800
         TabIndex        =   20
         Top             =   5040
         Visible         =   0   'False
         Width           =   4335
         _Version        =   851972
         _ExtentX        =   7646
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
      End
      Begin XtremeSuiteControls.FlatEdit txtIDNo 
         Height          =   375
         Left            =   -63280
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   4215
         _Version        =   851972
         _ExtentX        =   7435
         _ExtentY        =   661
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
      End
      Begin XtremeSuiteControls.RadioButton rBttn1 
         Height          =   255
         Left            =   -66640
         TabIndex        =   22
         Top             =   6480
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proof of billing like Credit Cards, etc."
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rBttn2 
         Height          =   255
         Left            =   -66640
         TabIndex        =   23
         Top             =   6240
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Bank Statement not older than 3 months"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rBttn3 
         Height          =   255
         Left            =   -66640
         TabIndex        =   24
         Top             =   6720
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851972
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registered mail not older 2 months"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Left            =   3600
         TabIndex        =   41
         Top             =   600
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "DATE APPLIED:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label17 
         Height          =   375
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "MEMBER ID:"
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
      Begin XtremeSuiteControls.Label Label15 
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   4680
         Width           =   2295
         _Version        =   851972
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "PERMANENT ADDRESS:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label18 
         Height          =   255
         Left            =   8880
         TabIndex        =   38
         Top             =   4680
         Width           =   1095
         _Version        =   851972
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ZIP CODE:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   5160
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "CURRENT ADDRESS:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Left            =   8760
         TabIndex        =   36
         Top             =   5160
         Width           =   1215
         _Version        =   851972
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ZIP CODE:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label13 
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   5760
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "CELLPHONE NUMBER:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   6240
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "E-MAIL ADDRESS:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label25 
         Height          =   255
         Left            =   -67720
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   4695
         _Version        =   851972
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "EXTENSION (Immediate Family Members only)"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label24 
         Height          =   255
         Left            =   -67720
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RELATION TO PRINCIPAL MEMBER:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label27 
         Height          =   255
         Left            =   -67840
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
         _Version        =   851972
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "CURRENT ADDRESS:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label26 
         Height          =   255
         Left            =   -60760
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851972
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ZIP CODE:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label11 
         Height          =   255
         Left            =   -67600
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "CELLPHONE NUMBER:"
         BackColor       =   14737632
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
         Height          =   255
         Left            =   -62080
         TabIndex        =   28
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851972
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "E-MAIL ADDRESS:"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label19 
         Height          =   255
         Left            =   -67600
         TabIndex        =   27
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851972
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "PRIMARY ID PRESENTED"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label23 
         Height          =   255
         Left            =   -62080
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "I.D./ Card Number"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label20 
         Height          =   255
         Left            =   -66880
         TabIndex        =   25
         Top             =   5880
         Visible         =   0   'False
         Width           =   5655
         _Version        =   851972
         _ExtentX        =   9975
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "SECONDARY FORM OF IDENTIFICATION"
         BackColor       =   14737632
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
Attribute VB_Name = "frmModifyMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat As String
Dim other As String
Private Sub Form_Load()
     tabMain(0).Caption = "Personal Information"
     tabMain(1).Caption = "Other Information Information"
     setFields
     getCount
     clear
     tabMain.SelectedItem = 0
End Sub

Private Sub txtFName_GotFocus()
    If txtFName.Text = "First Name" Then txtFName.Text = ""
End Sub

Private Sub txtFName_LostFocus()
    If txtFName.Text = "" Then txtFName.Text = "First Name"
End Sub

Private Sub txtLName_GotFocus()
    If txtLName.Text = "Last Name" Then txtLName.Text = ""
End Sub

Private Sub txtLName_LostFocus()
    If txtLName.Text = "" Then txtLName.Text = "Last Name"
End Sub

Private Sub txtMName_GotFocus()
    If txtMName.Text = "Middle Name" Then txtMName.Text = ""
End Sub

Private Sub txtMName_LostFocus()
    If txtMName.Text = "" Then txtMName.Text = "Middle Name"
End Sub

Private Sub txtexFName_GotFocus()
    txtExFName.Text = ""
End Sub

Private Sub txtexFName_LostFocus()
    If txtExFName.Text = "" Then txtExFName.Text = "First Name"
End Sub

Private Sub txtexLName_GotFocus()
    txtExLName.Text = ""
End Sub

Private Sub txtexLName_LostFocus()
    If txtExLName.Text = "" Then txtExLName.Text = "Last Name"
End Sub

Private Sub txtexMName_GotFocus()
    txtExMName.Text = ""
End Sub

Private Sub txtexMName_LostFocus()
    If txtExMName.Text = "" Then txtExMName.Text = "Last Name"
End Sub

Private Sub cboID_GotFocus()
    cboID.Text = ""
End Sub

Private Sub cboID_LostFocus()
    If cboID.Text = "Type of ID" Or cboID.Text = "" Then cboID.Text = "Type of ID"
End Sub
Private Sub bttnBrowse_Click()
    CommonDialog1.Filter = "Picture Files(*.jpg,*gif)|*.jpg;*.gif"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)

    Image1.Height = 2295
    Image1.Width = 3735
    Image1.Stretch = True
End Sub

Sub setFields()
With Me
    .cboGender.AddItem "Male"
    .cboGender.AddItem "Female"
    
    .cboID.AddItem "Passport"
    .cboID.AddItem "Student ID"
    .cboID.AddItem "PRC License ID"
    .cboID.AddItem "Driver's License"
    .cboID.AddItem "T.I.N."
    .cboID.AddItem "Postal"
    .cboID.AddItem "Company"
    
End With
End Sub

'
Sub getCount()

Dim ctr As Integer
    ctr = 1
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select MemberID from tbl_Members", cn, 1, 2
        Do Until .EOF
            ctr = ctr + 1
            .MoveNext
        Loop
        Me.txtMemberID.Text = ctr
    End With
    Set rs = Nothing
End Sub

Private Sub bttnCancel_Click()
    Unload Me
End Sub

Sub clear()
With Me

    .txtExLName.Text = "Last Name"
    .txtExFName.Text = "First Name"
    .txtExMName.Text = "Middle Name"
    .cboID.Text = "Type of ID"
    
    .rBttnMarried.value = False
    .rBttnSingle.value = False
    .rBttnWidow.value = False
    
    rBttn1.value = False
    rBttn2.value = False
    rBttn3.value = False
    
    .txtLName.Text = "Last Name"
    .txtFName.Text = "First Name"
    .txtMName.Text = "Middle Name"

    .cboGender.Text = ""
    .txtPerAdd.Text = ""
    .txtPerZip.Text = ""
    .txtCurAdd.Text = ""
    .txtCurZip.Text = ""
    .txtCelNo.Text = ""
    .txtEmailAdd.Text = ""
    .txtAge = ""
    .txtRelation = ""
    .txtExCurAdd = ""
    .txtExCurZip.Text = ""
    .txtExCelNo.Text = ""
    .txtExEmailAdd.Text = ""
    .txtIDNo.Text = ""
    
End With
End Sub

Sub proceedToEdit()

On Error GoTo err

    Dim byteData() As Byte
    
    On Error GoTo err
    If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Binary As #1
    ReDim byteData(FileLen(CommonDialog1.FileName))
    Get #1, , byteData
    Close #1
    End If
Set rs = New ADODB.Recordset
    With rs
       .Open "Select * from tbl_Members where memberID like '" & Me.txtMemberID.Text & "'", cn, 1, 2
        
       ' !MemberID = Me.txtMemberID.Text
        !Lname = Me.txtLName.Text
        !Fname = Me.txtFName.Text
        !Mname = Me.txtMName.Text
        !Sex = Me.cboGender.Text
        !Age = Me.txtAge.Text
        !CivilStatus = stat
        !Bdate = Me.dtBdate.value
        !CurAdd = Me.txtCurAdd.Text
        !CurZIP = Me.txtCurZip.Text
        !PerAdd = Me.txtPerAdd.Text
        !PerZIP = Me.txtPerZip.Text
        !Cellphone = Me.txtCelNo.Text
        !Email = Me.txtEmailAdd.Text
        !IDPresented = Me.cboID.Text
        !IDCardNo = Me.txtIDNo.Text
        !OtherID = other
        !ExLname = Me.txtExLName.Text
        !ExFname = Me.txtExFName.Text
        !ExMname = Me.txtExMName.Text
        !Sex = Me.cboGender.Text
        !Relation = Me.txtRelation.Text
        !ExCurAdd = Me.txtExCurAdd.Text
        !ExZIP = Me.txtExCurZip.Text
        !ExCellPhone = Me.txtExCelNo.Text
        !ExEmail = Me.txtExEmailAdd.Text
        !Status = "ACTIVE"
         !TransID = ctr
      
        !DateApplied = Me.dtpApplication.value
          If CommonDialog1.FileName <> "" Then !pic.AppendChunk byteData
        .Update
        wDone = "Edit New Member" & "/" & "Member ID" & " - " & Me.txtMemberID.Text
        Call auditT
            
        If MsgBox("Record Successfully Updated.", vbInformation + vbOKOnly) = vbOK Then Call clear
            tabMain.SelectedItem = 0
            getCount
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Sub msg()
    MsgBox "Missing Field(s). Please completely fill up.", vbInformation + vbOKOnly
End Sub

Sub chckFields()
'
' If Text_Empty(Me.txtLName) = True Then Exit Sub
' If Text_Empty(Me.txtFName) = True Then Exit Sub
' If Text_Empty(Me.txtFName) = True Then Exit Sub
'If Text_Empty(Me.txtAge) = True Then Exit Sub
' If Text_Empty(Me.txtCurAdd) = True Then Exit Sub
' If Text_Empty(Me.txtCurZip) = True Then Exit Sub
' If Text_Empty(Me.txtPerAdd) = True Then Exit Sub
'If Text_Empty(Me.txtPerZip) = True Then Exit Sub
'If Text_Empty(Me.txtCelNo) = True Then Exit Sub
'If Text_Empty(Me.txtEmailAdd) = True Then Exit Sub
' If Text_Empty(Me.txtExLName) = True Then Exit Sub
' If Text_Empty(Me.txtExFName) = True Then Exit Sub
' If Text_Empty(Me.txtExMName) = True Then Exit Sub
' If Text_Empty(Me.txtRelation) = True Then Exit Sub
' If Text_Empty(Me.txtExCurAdd) = True Then Exit Sub
' If Text_Empty(Me.txtExCurZip) = True Then Exit Sub
' If Text_Empty(Me.txtIDNo) = True Then Exit Sub
If txtLName.Text = "" Or txtLName.Text = "Last Name" Then
    Call msg
    tabMain.SelectedItem = 0
    txtLName.SetFocus
    Exit Sub
ElseIf txtFName.Text = "" Or txtFName.Text = "First Name" Then
    Call msg
    tabMain.SelectedItem = 0
    txtFName.SetFocus
    Exit Sub
ElseIf txtMName.Text = "" Or txtMName.Text = "Middle Name" Then
    Call msg
    tabMain.SelectedItem = 0
    txtMName.SetFocus
    Exit Sub
ElseIf txtAge.Text = "" Then
    Call msg
    tabMain.SelectedItem = 0
    txtAge.SetFocus
    Exit Sub
ElseIf cboGender.Text = "" Then
    Call msg
    tabMain.SelectedItem = 0
    cboGender.SetFocus
    Exit Sub
ElseIf rBttnSingle.value = False And rBttnMarried.value = False And rBttnWidow.value = False Then
    tabMain.SelectedItem = 0
    MsgBox "What is your Civil Status?", vbQuestion + vbOKOnly
    Exit Sub
ElseIf txtPerAdd.Text = "" Then
    Call msg
    tabMain.SelectedItem = 0
    txtPerAdd.SetFocus
    Exit Sub
ElseIf txtCurAdd.Text = "" Then
    Call msg
    tabMain.SelectedItem = 0
    txtCurAdd.SetFocus
    Exit Sub
ElseIf txtExLName.Text = "" Or txtExLName.Text = "Last Name" Then
    Call msg
    tabMain.SelectedItem = 1
    txtExLName.SetFocus
    Exit Sub
ElseIf txtExFName.Text = "" Or txtExFName.Text = "First Name" Then
    Call msg
    tabMain.SelectedItem = 1
    txtExFName.SetFocus
    Exit Sub
ElseIf txtExMName.Text = "" Or txtExMName.Text = "Middle Name" Then
    Call msg
    tabMain.SelectedItem = 1
    txtExMName.SetFocus
    Exit Sub
ElseIf txtRelation.Text = "" Then
    Call msg
    tabMain.SelectedItem = 1
    txtRelation.SetFocus
    Exit Sub
ElseIf txtExCurAdd.Text = "" Then
    Call msg
    tabMain.SelectedItem = 1
    txtExCurAdd.SetFocus
    Exit Sub
ElseIf cboID.Text = "" Then
    Call msg
    tabMain.SelectedItem = 1
    cboID.SetFocus
    Exit Sub
    tabMain.SelectedItem = 1
ElseIf txtIDNo.Text = "" Then
    Call msg
    tabMain.SelectedItem = 1
    txtIDNo.SetFocus
    Exit Sub
    
    ElseIf cboID.Text = "" Or cboID.Text = "Type of ID" Then
    Call msg
    tabMain.SelectedItem = 1
    cboID.SetFocus
    Exit Sub
ElseIf rBttn1.value = False And rBttn2.value = False And rBttn3.value = False Then
    tabMain.SelectedItem = 1
    MsgBox "No Secondary Identification", vbInformation + vbOKOnly
    Exit Sub

End If

End Sub
'
Sub assgnValue()

    If rBttnSingle.value = True Then stat = rBttnSingle.Caption
    If rBttnMarried.value = True Then stat = rBttnMarried.Caption
    If rBttnWidow.value = True Then stat = rBttnWidow.Caption



    If rBttn1.value = True Then other = rBttn1.Caption
    If rBttn2.value = True Then other = rBttn2.Caption
    If rBttn3.value = True Then other = rBttn3.Caption

End Sub

Private Sub bttnUpdate_Click()
    Call assgnValue
    Call chckFields
    
    Call proceedToEdit
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub dtBdate_Change()
Dim iyears As Integer
    Dim tmonth As Integer
    Dim Bmonth As Integer

    iyears = DateDiff("yyyy", "01/01/" & DatePart("yyyy", dtBdate.value), _
    "01/01/" & DatePart("yyyy", Now)) - 1

    tmonth = DatePart("m", Now)
    Bmonth = DatePart("m", dtBdate.value)

    Select Case tmonth
        Case Is > Bmonth
            iyears = iyears + 1
        Case Is = Bmonth
            If DatePart("d", dtBdate.value) <= _
            DatePart("d", Now) Then iyears = iyears + 1

    End Select
    Me.txtAge.Text = iyears
End Sub







