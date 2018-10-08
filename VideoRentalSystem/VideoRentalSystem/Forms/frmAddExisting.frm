VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddExisting 
   Caption         =   "New CD Entry"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   12600
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   12495
      _Version        =   851972
      _ExtentX        =   22040
      _ExtentY        =   4895
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
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   720
      ScaleHeight     =   2355
      ScaleWidth      =   3795
      TabIndex        =   34
      Top             =   480
      Width           =   3855
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   3735
      End
   End
   Begin XtremeSuiteControls.PushButton bttnAddExisting 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   7080
      Width           =   2055
      _Version        =   851972
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Add Existing"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmAddExisting.frx":0000
   End
   Begin XtremeSuiteControls.PushButton bttnSave 
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
      _Version        =   851972
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmAddExisting.frx":039A
   End
   Begin XtremeSuiteControls.PushButton bttnCancel 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   7080
      Width           =   2055
      _Version        =   851972
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmAddExisting.frx":0734
   End
   Begin XtremeSuiteControls.FlatEdit txtDVDNo 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   240
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
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
   Begin XtremeSuiteControls.ListBox lstID 
      Height          =   300
      Left            =   6960
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   529
      _StockProps     =   77
      ForeColor       =   12582912
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
   End
   Begin XtremeSuiteControls.FlatEdit txtTitle 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   960
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
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
      Left            =   6960
      TabIndex        =   13
      Top             =   1440
      Width           =   4335
      _Version        =   851972
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.ComboBox cboGenre 
      Height          =   315
      Left            =   6960
      TabIndex        =   14
      Top             =   1440
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtActor 
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   1920
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
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
   Begin XtremeSuiteControls.FlatEdit txtDirector 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   2400
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
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
   Begin XtremeSuiteControls.FlatEdit txtYearRealese 
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   2880
      Width           =   4335
      _Version        =   851972
      _ExtentX        =   7646
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboRelease 
      Height          =   315
      Left            =   6960
      TabIndex        =   21
      Top             =   2880
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCategory 
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   3360
      Width           =   4335
      _Version        =   851972
      _ExtentX        =   7646
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboCategory 
      Height          =   315
      Left            =   6960
      TabIndex        =   23
      Top             =   3360
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCDType 
      Height          =   315
      Left            =   6960
      TabIndex        =   24
      Top             =   3840
      Width           =   4335
      _Version        =   851972
      _ExtentX        =   7646
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboCDType 
      Height          =   315
      Left            =   6960
      TabIndex        =   25
      Top             =   3840
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtQuantity 
      Height          =   495
      Left            =   1560
      TabIndex        =   30
      Top             =   3720
      Width           =   3015
      _Version        =   851972
      _ExtentX        =   5318
      _ExtentY        =   873
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
   Begin XtremeSuiteControls.ListBox lstCategory 
      Height          =   615
      Left            =   6840
      TabIndex        =   31
      Top             =   6480
      Visible         =   0   'False
      Width           =   4575
      _Version        =   851972
      _ExtentX        =   8070
      _ExtentY        =   1085
      _StockProps     =   77
      ForeColor       =   12582912
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton bttnCategory 
      Height          =   375
      Left            =   11040
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
      _Version        =   851972
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   12582912
      Appearance      =   6
      PushButtonStyle =   1
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   2655
      Left            =   600
      TabIndex        =   33
      Top             =   360
      Width           =   4095
      _Version        =   851972
      _ExtentX        =   7223
      _ExtentY        =   4683
      _StockProps     =   79
      ForeColor       =   12582912
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton bttnBrowse 
      Height          =   495
      Left            =   600
      TabIndex        =   35
      Top             =   3120
      Width           =   4095
      _Version        =   851972
      _ExtentX        =   7223
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Browse Image"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   3720
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Quantity:"
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
      Left            =   5880
      TabIndex        =   28
      Top             =   3840
      Width           =   975
      _Version        =   851972
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CD Type"
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
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Category"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   2880
      Width           =   1575
      _Version        =   851972
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Year Realese"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   2400
      Width           =   855
      _Version        =   851972
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Director"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   1920
      Width           =   1575
      _Version        =   851972
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actor/Actress"
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
      Left            =   6000
      TabIndex        =   17
      Top             =   1440
      Width           =   855
      _Version        =   851972
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Genre"
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
      Left            =   6240
      TabIndex        =   12
      Top             =   960
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Title"
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
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   240
      Width           =   975
      _Version        =   851972
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "ID No:"
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
   Begin XtremeSuiteControls.Label lblTime 
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   7440
      Width           =   2175
      _Version        =   851972
      _ExtentX        =   3836
      _ExtentY        =   873
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   7560
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Time"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblDate 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   7560
      Width           =   2655
      _Version        =   851972
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAddExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nwEntry As Boolean

Private Sub bttnAddExisting_Click()
    nwEntry = False
    Me.txtDVDNo.Enabled = True
    Me.txtQuantity.Enabled = True
    Me.bttnCancel.Caption = "Cancel"
    Me.bttnBrowse.Enabled = True

End Sub

Private Sub bttnBrowse_Click()
    CommonDialog1.Filter = "Picture Files(*.jpg,*.gif,*.bmp)|*.jpg;*.gif;*.bmp"
    CommonDialog1.ShowOpen
    Image1.Picture = LoadPicture(CommonDialog1.FileName)

    Image1.Height = 2295
    Image1.Width = 3735
    Image1.Stretch = True
End Sub

Private Sub bttnCancel_Click()
    If Me.bttnCancel.Caption = "Close" Then
        Unload Me
    Else
    Call lockF
    Call clearF
        Me.bttnCancel.Caption = "Close"
        Me.bttnSave.Enabled = False
        Me.bttnAddExisting.Enabled = True

    End If
End Sub

Private Sub bttnCategory_Click()
    If Me.lstCategory.Visible = False Then
    Me.lstCategory.Visible = True
End If
End Sub



Private Sub bttnSave_Click()
On Error GoTo err
If nwEntry = True Then
    Dim byteData() As Byte
    Open CommonDialog1.FileName For Binary As #1
    ReDim byteData(FileLen(CommonDialog1.FileName))
    Get #1, , byteData
    Close #1
    
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_CDList", cn, 1, 2
        .AddNew
        !CD_ID = txtDVDNo.Text
        !Title = txtTitle.Text
        !Ggenre = txtGenre.Text
        !director = txtDirector.Text
        !actor = txtActor.Text
        !YearRelease = txtYearRealese.Text
        !CDType = txtCDType.Text
        !goodFor = "1"
        !Category = txtCategory.Text
        !pic.AppendChunk byteData
        !dateAcquired = Me.lblDate.Caption
        !Status = "AVAILABLE"
        .Update
    End With
    
    Call getCount
    
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_IndividualCDList", cn, 1, 2
        For I = 1 To CInt(Me.txtQuantity.Text)
            .AddNew
            !Ino = ctr
            !CD_ID = Me.txtDVDNo.Text
            !dateAcquired = Date
            !Qnty = "1"
            !Status = "AVAILABLE"
            .Update
            
            ctr = ctr + 1
        Next I
        .Close
    End With
        Call addGenre
        Call addDirector
        Call addActor
        
                    'audit transaction
            wDone = "New CD ENTRY" & "/" & "ID Number" & " - " & Me.txtDVDNo.Text
            Call auditT
            
            If MsgBox("New Entry has been Successfully Added.", vbInformation + vbOKOnly) = vbOK Then clearF
            Me.bttnAddExisting.Enabled = True
            Me.bttnSave.Enabled = False
            Me.bttnCancel.Caption = "Close"
            Call cboItems
            Call clearF
            Set rs = Nothing
            
    Else
        Call getCount
        Set rs = New ADODB.Recordset
         With rs
            .Open "Select * from tbl_IndividualCDList", cn, 1, 2
            For I = 1 To CInt(Me.txtQuantity.Text)
                .AddNew
                !Ino = ctr
                !CD_ID = Me.txtDVDNo.Text
                !dateAcquired = Date
                !Qnty = "1"
                !Status = "AVAILABLE"
                .Update
            
            ctr = ctr + 1
        Next I
        .Close
    End With
                '***********Audit Trail
            wDone = "EXISTING CD ENTRY" & "/" & "ID Number" & " - " & Me.txtDVDNo.Text
            Call auditT
            If MsgBox("New Entry has been Successfully Added.", vbInformation + vbOKOnly) = vbOK Then clearF
                Me.bttnAddExisting.Enabled = True
               
                Me.bttnBrowse.Enabled = False
                Me.bttnSave.Enabled = False
                Me.bttnCancel.Caption = "Close"
                Call cboItems
                Call clearF
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cboCategory_Click()
    Me.txtCategory.Locked = False
    Me.txtCategory.Text = cboCategory.Text
    Me.txtCategory.Locked = True
End Sub

Private Sub cboCDType_Click()
    Me.txtCDType.Locked = False
    Me.txtCDType.Text = cboCDType.Text
    Me.txtCDType.Locked = True
End Sub

Private Sub cboGenre_Click()
    Me.txtGenre.Locked = False
    If Me.txtGenre.Text = "" Then
        txtGenre.Text = txtGenre.Text & cboGenre.Text
        cboGenre.RemoveItem (cboGenre.ListIndex)
    Else
        txtGenre.Text = txtGenre.Text & "/" & cboGenre.Text
        cboGenre.RemoveItem (cboGenre.ListIndex)
    End If
Me.txtGenre.Locked = True
End Sub

Private Sub cboRelease_Click()
    Me.txtYearRealese.Locked = False
    Me.txtYearRealese.Text = cboRelease.Text
    Me.txtYearRealese.Locked = True
End Sub

Private Sub Form_Load()
    Me.lblDate.Caption = Date
    Me.lblTime.Caption = Time

    Me.bttnCancel.Caption = "Close"
    Me.bttnSave.Enabled = False
    Me.txtQuantity.Text = 0

    Me.Top = 4000
    Me.Left = 2300

    Call lockF
    Call cboItems
End Sub

Sub lockF()
With Me
    .txtActor.Enabled = False
    .txtCategory.Enabled = False
    .txtCDType.Enabled = False
    .txtDirector.Enabled = False
    .txtDVDNo.Enabled = False
    .txtGenre.Enabled = False
    .txtQuantity.Enabled = False
    .txtTitle.Enabled = False
    .txtYearRealese.Enabled = False
    .bttnBrowse.Enabled = False
    .cboCategory.Enabled = False
    .cboCDType.Enabled = False
    .cboGenre.Enabled = False
    .cboRelease.Enabled = False
End With
End Sub

Sub unlockF()
With Me
    .txtActor.Enabled = True
    .txtCategory.Enabled = True
    .txtCDType.Enabled = True
    .txtDirector.Enabled = True
    .txtDVDNo.Enabled = True
    .txtGenre.Enabled = True
    .txtQuantity.Enabled = True
    .txtTitle.Enabled = True
    .txtYearRealese.Enabled = True
    .bttnBrowse.Enabled = True
    .cboCategory.Enabled = True
    .cboCDType.Enabled = True
    .cboGenre.Enabled = True
    .cboRelease.Enabled = True
End With
End Sub

Sub clearF()
With Me
    .txtActor.Text = ""
    .txtCategory.Text = ""
    .txtCDType.Text = ""
    .txtDirector.Text = ""
    .txtDVDNo.Text = ""
    .txtGenre.Text = ""
    .txtQuantity.Text = ""
    .txtTitle.Text = ""
    .txtYearRealese.Text = ""
    Image1.Picture = LoadPicture("")
End With
End Sub

Sub newEntry()
    Dim x As String
    ctr = 1
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_CDList", cn, 1, 2
        x = Format(Date, "mm-dd-yyyy")
        If .EOF = True Then
            ctr = 1
            Me.txtDVDNo.Text = x & "-" & ctr
        Else
            Do Until .EOF
            ctr = ctr + 1
            .MoveNext
            Loop
            Me.txtDVDNo.Text = x & "-" & ctr
        End If
            .Close
    End With
End Sub

Sub cboItems()
    cboGenre.clear
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Genre", cn, 1, 2
        If .EOF = False Then
            Do Until .EOF
            cboGenre.AddItem (!genre)
            .MoveNext
            Loop
        End If
    End With
    
    Me.cboRelease.clear
    Dim ctr As Integer
    ctr = 2012
        Do Until ctr = 1970
        cboRelease.AddItem (ctr)
        If ctr = 1970 Then Exit Do
            ctr = ctr - 1
        Loop
    
    With Me
        .cboCDType.clear
        .cboCDType.AddItem "DVD"
        .cboCDType.AddItem "VCD"
    End With

    With Me
        .cboCategory.clear
        .cboCategory.AddItem "English"
        .cboCategory.AddItem "Tagalog"
    End With
End Sub


Sub addGenre()
    Dim genre1() As String
    Dim sPlitCtr As Integer

    Set rs = New ADODB.Recordset
    With rs
        genre1 = Split(Me.txtGenre.Text, "/")
        For sPlitCtr = 0 To UBound(genre1)

        .Open "Select * from tbl_Genre", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !genre = genre1(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If UCase(genre1(sPlitCtr)) = UCase(!genre) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !genre = genre1(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub

Sub addDirector()
Dim director() As String
Dim sPlitCtr As Integer

Set rs = New ADODB.Recordset
    With rs
        director = Split(Me.txtDirector.Text, "/")
        For sPlitCtr = 0 To UBound(director)

        .Open "Select * from tbl_Director", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !DirectorName = director(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If Trim(UCase(director(sPlitCtr))) = Trim(UCase(!DirectorName)) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !DirectorName = director(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub

Sub addActor()
Dim actor() As String
Dim sPlitCtr As Integer

Set rs = New ADODB.Recordset
    With rs
        actor = Split(Me.txtActor.Text, "/")
        For sPlitCtr = 0 To UBound(actor)

        .Open "Select * from tbl_Actor_Actress", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !ActorActressName = actor(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If Trim(UCase(actor(sPlitCtr))) = Trim(UCase(!ActorActressName)) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !ActorActressName = actor(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub

Private Sub lstID_Click()
ndex = Me.lstID.ListIndex
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CDList where CD_ID like '" & Me.lstID.List(ndex) & "'", cn, 1, 2
    
    Me.txtDVDNo.Text = Me.lstID.List(ndex)
    Me.txtTitle.Text = !Title
    Me.txtGenre.Text = !Ggenre
    Me.txtActor.Text = !actor
    Me.txtDirector.Text = !director
    Me.txtYearRealese.Text = !YearRelease
    Me.txtCategory.Text = !Category
    Me.txtCDType.Text = !CDType
    
    Set Me.Image1.DataSource = rs
    Me.Image1.DataField = "pic"
    Image1.Height = 2295
    Image1.Width = 3735
    Image1.Stretch = True
    .Close
End With
    Me.lstID.Visible = False
End Sub

Private Sub txtDVDNo_Change()
If nwEntry = False Then
ndex = 0
    lstID.clear
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_CDList where  CD_ID like '" & Me.txtDVDNo.Text & "%'", cn, 1, 2
        
        Do Until .EOF
            lstID.AddItem !CD_ID
            If ndex >= 2580 Then
            Else
            ndex = ndex + 300
            End If
            .MoveNext
        Loop
        .Close
        lstID.Visible = True
        lstID.Height = ndex
        If Me.txtDVDNo.Text = "" Then lstID.Visible = False
    End With
End If
End Sub

Private Sub txtQuantity_Change()
With Me.ListView1
    .ListItems.clear
    .ColumnHeaders.clear
    
    .ColumnHeaders.Add , , "Individual No", 1500
    .ColumnHeaders.Add , , "CD Identification", 1500
    .ColumnHeaders.Add , , "Title", 1500
    .ColumnHeaders.Add , , "Genre", 1500
    .ColumnHeaders.Add , , "Actor/Actress", 1500
    .ColumnHeaders.Add , , "Director", 1500
    .ColumnHeaders.Add , , "Year Realese", 1500
    .ColumnHeaders.Add , , "Category", 1500
    .ColumnHeaders.Add , , "CD Type", 1500
End With

Call getCount
If Me.txtQuantity.Text = "" Then Me.txtQuantity.Text = "0"
    For I = 1 To CInt(Me.txtQuantity.Text)
            Set lst = Me.ListView1.ListItems.Add(, , ctr)
        lst.ListSubItems.Add , , Me.txtDVDNo.Text
        lst.ListSubItems.Add , , Me.txtTitle.Text
        lst.ListSubItems.Add , , Me.txtGenre.Text
        lst.ListSubItems.Add , , Me.txtActor.Text
        lst.ListSubItems.Add , , Me.txtDirector.Text
        lst.ListSubItems.Add , , Me.txtYearRealese.Text
        lst.ListSubItems.Add , , Me.txtCategory.Text
        lst.ListSubItems.Add , , Me.txtCDType.Text
        ctr = ctr + 1
    Next I


If Me.txtQuantity.Text > 0 Then Me.bttnSave.Enabled = True

End Sub

Sub getCount()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_IndividualCDList", cn, 1, 2
    ctr = 1
    If .EOF = True Then
        ctr = 1
    Else
        Do Until .EOF
            ctr = ctr + 1
            .MoveNext
        Loop
    End If
    .Close
End With
End Sub


