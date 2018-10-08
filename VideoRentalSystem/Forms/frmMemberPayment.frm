VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmMemberPayment 
   Caption         =   "Membership Payment"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _Version        =   851972
      _ExtentX        =   10398
      _ExtentY        =   8070
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      Begin XtremeSuiteControls.FlatEdit txtMembershipFee 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
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
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPayment 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   2160
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
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
      End
      Begin XtremeSuiteControls.FlatEdit txtChange 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   2880
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
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
         Enabled         =   0   'False
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnRegister 
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   3840
         Width           =   2055
         _Version        =   851972
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Register"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmMemberPayment.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnCancel 
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   3840
         Width           =   2295
         _Version        =   851972
         _ExtentX        =   4048
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
         Picture         =   "frmMemberPayment.frx":039A
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
         _Version        =   851972
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Membership Fee:"
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
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
         _Version        =   851972
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Payment:"
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
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   2880
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Change:"
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
      Begin XtremeSuiteControls.Label Label17 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   3015
         _Version        =   851972
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Payment for Membership"
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
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMemberPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRegister_Click()
    Unload Me
    Call frmNewMember.proceedToAdd
    Call getIncomeID
  
    'for income
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Income", cn, 1, 2
            .AddNew
            !incomeID = incomeID
            !amount = Me.txtMembershipFee.Text
            !iDate = Now
            !iType = "MEMBERSHIP FEE"
            !UserID = user
        .Update
        .Close
End With
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Membership", cn, 1, 2
         Me.txtMembershipFee.Text = !amount
    .Close
End With
Me.btnRegister.Enabled = False
End Sub

Private Sub txtPayment_Change()

On Error Resume Next
    If CInt(Me.txtPayment.Text) >= CInt(Me.txtMembershipFee.Text) Then
        Me.txtChange.Text = CInt(Me.txtPayment.Text) - CInt(Me.txtMembershipFee.Text)
        Me.btnRegister.Enabled = True
    Else
        Me.btnRegister.Enabled = False
    End If

    If Me.txtPayment.Text = "" Then
        Me.btnRegister.Enabled = False
        Me.txtChange.Text = ""
    End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    Else
    KeyAscii = 0
    End If
End Sub

