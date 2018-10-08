VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmDVDInfo 
   Caption         =   "CD/DVD Info"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   3000
      ScaleHeight     =   2355
      ScaleWidth      =   3795
      TabIndex        =   21
      Top             =   600
      Width           =   3855
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   2655
      Left            =   2760
      TabIndex        =   20
      Top             =   480
      Width           =   4335
      _Version        =   851972
      _ExtentX        =   7646
      _ExtentY        =   4683
      _StockProps     =   79
      ForeColor       =   12582912
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   615
      Left            =   3840
      TabIndex        =   22
      Top             =   6960
      Width           =   3135
      _Version        =   851972
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmDVDInfo.frx":0000
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   -960
      Picture         =   "frmDVDInfo.frx":039A
      Top             =   7560
      Width           =   11535
   End
   Begin VB.Image Image6 
      Height          =   420
      Left            =   0
      Picture         =   "frmDVDInfo.frx":0874
      Top             =   0
      Width           =   11535
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   495
      Left            =   5280
      TabIndex        =   19
      Top             =   6120
      Width           =   1695
      _Version        =   851972
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Available Copies:"
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
   Begin XtremeSuiteControls.Label Label9 
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   5400
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Number of Copies:"
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
   Begin XtremeSuiteControls.Label Label8 
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   4680
      Width           =   975
      _Version        =   851972
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Category:"
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
   Begin XtremeSuiteControls.Label Label7 
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "CD Type:"
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
   Begin XtremeSuiteControls.Label Label6 
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
      _Version        =   851972
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Year Release:"
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   6120
      Width           =   735
      _Version        =   851972
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Director:"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   5400
      Width           =   1335
      _Version        =   851972
      _ExtentX        =   2355
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Actor/Actress:"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   4680
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Genre:"
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
   Begin XtremeSuiteControls.Label lblAvaiCopy 
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   6120
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblNoCopy 
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   5400
      Width           =   3615
      _Version        =   851972
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.Label lblCategory 
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   4680
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblCdType 
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   4080
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblYrRelease 
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblDirector 
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   6120
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblActor 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   5400
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblGenre 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   4680
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.Label lblTitle 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
      _Version        =   851972
      _ExtentX        =   5953
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
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Title:"
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
   Begin XtremeSuiteControls.Label lblID 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
      _Version        =   851972
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.Label lbli 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "ID Number:"
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
End
Attribute VB_Name = "frmDVDInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnClose_Click()
    Me.Hide
End Sub

