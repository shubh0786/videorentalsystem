VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmAddPenalty 
   Caption         =   "Add Penalty"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton bttnCancel 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
      _Version        =   851972
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Picture         =   "frmAddPenalty.frx":0000
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _Version        =   851972
      _ExtentX        =   7435
      _ExtentY        =   5953
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   3735
         _Version        =   851972
         _ExtentX        =   6588
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton bttnOk 
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1335
            _Version        =   851972
            _ExtentX        =   2355
            _ExtentY        =   661
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
            Picture         =   "frmAddPenalty.frx":039A
         End
      End
      Begin XtremeSuiteControls.RadioButton rBttn 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   1560
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Lose"
         BackColor       =   0
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
      End
      Begin XtremeSuiteControls.RadioButton rBttn 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Damage"
         BackColor       =   -2147483628
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
      End
      Begin XtremeSuiteControls.RadioButton rBttn 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Tampering"
         BackColor       =   -2147483628
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
      End
   End
End
Attribute VB_Name = "frmAddPenalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tampered, lose, damage As Boolean
Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnOk_Click()
    oTampering = False
    oLose = False
    oDamage = False
If rBttn(0).value = True Then
    oLose = True
ElseIf rBttn(1).value = True Then
    oDamage = True
ElseIf rBttn(2).value = True Then
    oTampering = True
End If
    Unload Me
If bReturn = True Then
    Call frmReturn.calcAdditional
    frmReturn.bttnAdditional.Enabled = False
End If
End Sub


