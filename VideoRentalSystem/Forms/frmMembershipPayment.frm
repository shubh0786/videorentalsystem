VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmMembershipPaymentSettings 
   Caption         =   "Membership Payment Settings"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton bttnSave 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Save"
      ForeColor       =   0
      BackColor       =   -2147483641
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
      Picture         =   "frmMembershipPayment.frx":0000
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   5175
      _Version        =   851972
      _ExtentX        =   9128
      _ExtentY        =   3413
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
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.FixedTabWidth=   310
      ItemCount       =   1
      Item(0).Caption =   "Membership"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "Label11"
      Item(0).Control(1)=   "txtAmount"
      Begin XtremeSuiteControls.FlatEdit txtAmount 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   720
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
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   1095
         _Version        =   851972
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Amount:"
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
   Begin XtremeSuiteControls.PushButton bttnEdit 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Edit"
      ForeColor       =   0
      BackColor       =   -2147483641
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
      Picture         =   "frmMembershipPayment.frx":039A
   End
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Close"
      ForeColor       =   0
      BackColor       =   -2147483641
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
      Picture         =   "frmMembershipPayment.frx":0734
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5175
      _Version        =   851972
      _ExtentX        =   9128
      _ExtentY        =   6376
      _StockProps     =   79
      Appearance      =   6
   End
End
Attribute VB_Name = "frmMembershipPaymentSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Me.bttnEdit.Enabled = False
    Me.txtAmount.Enabled = True
    Me.bttnSave.Enabled = True
    Me.txtAmount.SetFocus
End Sub

Private Sub bttnSave_Click()
If Me.bttnSave.Caption = "Save" Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Membership", cn, 1, 2
        
        If .EOF = True Then
            .AddNew
            !amount = Me.txtAmount.Text
            .Update
        Else
            !amount = Me.txtAmount.Text
            .Update
        End If
        .Close
        MsgBox "Record Save.", vbInformation + vbOKOnly
    End With
    Me.bttnSave.Enabled = False
    Me.bttnEdit.Enabled = True
ElseIf Me.bttnSave.Caption = "New" Then
    Me.bttnSave.Caption = "Save"
    Me.txtAmount.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Membership", cn, 1, 2
    
    If .EOF = True Then
        Me.bttnSave.Caption = "New"
        Me.bttnEdit.Enabled = False
        Me.bttnSave.Enabled = True
    Else
        Me.txtAmount.Text = !amount
    
        Me.bttnSave.Caption = "Save"
        Me.bttnEdit.Enabled = True
    End If
    .Close
End With
End Sub

