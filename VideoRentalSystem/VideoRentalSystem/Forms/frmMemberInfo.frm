VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmMemberInfo 
   Caption         =   "Member's Info"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   3360
      ScaleHeight     =   2355
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   720
      Width           =   3855
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   2895
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _Version        =   851972
      _ExtentX        =   7223
      _ExtentY        =   5106
      _StockProps     =   79
      ForeColor       =   12582912
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   615
      Left            =   3840
      TabIndex        =   2
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
      Picture         =   "frmMemberInfo.frx":0000
   End
   Begin VB.Label Label14 
      Caption         =   "Address:"
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
      Left            =   6120
      TabIndex        =   30
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Relationship:"
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
      Left            =   5760
      TabIndex        =   29
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Extension Name:"
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
      Left            =   5400
      TabIndex        =   28
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "ID Card No.:"
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
      Left            =   5640
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "ID Presented:"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Email Address:"
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
      Left            =   5640
      TabIndex        =   25
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Address:"
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
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Civil Status:"
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
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Age:"
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
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Gender:"
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
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblExCurAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblExRelation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label lblExName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label lblIDPres 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label lblCardNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblEMail 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblCelNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Celphone No.:"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label lblCivil 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label lblAge 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label lblGender 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label lblBirth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Birthdate:"
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
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
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
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label label1 
      Caption         =   "Member ID:"
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
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   -720
      Picture         =   "frmMemberInfo.frx":039A
      Top             =   7560
      Width           =   11535
   End
   Begin VB.Image Image6 
      Height          =   420
      Left            =   -240
      Picture         =   "frmMemberInfo.frx":0874
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmMemberInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Image1.Width = Me.Width
End Sub

Private Sub bttnClose_Click()
    Me.Hide
End Sub

