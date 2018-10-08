VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmRentalRate 
   Caption         =   "Rental Rate"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
      _Version        =   851972
      _ExtentX        =   2778
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
      Picture         =   "frmRentalRate.frx":0000
   End
   Begin XtremeSuiteControls.PushButton bttnEdit 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
      _Version        =   851972
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Picture         =   "frmRentalRate.frx":039A
   End
   Begin XtremeSuiteControls.PushButton bttnSave 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
      _Version        =   851972
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Picture         =   "frmRentalRate.frx":0734
   End
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      _Version        =   851972
      _ExtentX        =   9763
      _ExtentY        =   5106
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
      Appearance      =   7
      Color           =   32
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "txtDVDRate"
      Item(0).Control(1)=   "Label17"
      Item(1).Caption =   "Item"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "txtVCDRate"
      Item(1).Control(1)=   "Label1"
      Begin XtremeSuiteControls.FlatEdit txtDVDRate 
         Height          =   495
         Left            =   -69280
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   4335
         _Version        =   851972
         _ExtentX        =   7646
         _ExtentY        =   873
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVCDRate 
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
         _Version        =   851972
         _ExtentX        =   7646
         _ExtentY        =   873
         _StockProps     =   77
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
         _Version        =   851972
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "CD Rental Rate"
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
      Begin XtremeSuiteControls.Label Label17 
         Height          =   375
         Left            =   -68560
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         _Version        =   851972
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "DVD Rental Rate"
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _Version        =   851972
      _ExtentX        =   10610
      _ExtentY        =   8070
      _StockProps     =   79
      Appearance      =   6
   End
End
Attribute VB_Name = "frmRentalRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        tabMain(0).Caption = "DVD"
        tabMain(1).Caption = "VCD"
        tabMain.SelectedItem = 0
        Call loadRecords
        Call lockF
        Me.bttnSave.Enabled = False
End Sub
Sub loadRecords()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_RentalRate", cn, 1, 2
    
    If .EOF = True Then
        Me.txtDVDRate.Text = "0.00"
        Me.txtVCDRate.Text = "0.00"
        Me.bttnEdit.Caption = "&New"
        
        tType = "Add Rental Rate"
    Else
        Me.bttnEdit.Caption = "&Edit"
        tType = "Edit Rental Rate"
        Do Until .EOF
            Me.txtDVDRate = !DVDAmount
            Me.txtVCDRate = !VCDAmount
            .MoveNext
        Loop
    End If
    .Close
End With
End Sub


Sub lockF()
    Me.txtDVDRate.Enabled = False
    Me.txtVCDRate.Enabled = False
End Sub

Sub unlockF()
    Me.txtDVDRate.Enabled = True
    Me.txtVCDRate.Enabled = True
End Sub

Private Sub txtDVDRate_GotFocus()
If Me.txtDVDRate.Text = "" Or Me.txtDVDRate.Text = "0.00" Then Me.txtDVDRate.Text = ""
End Sub

Private Sub txtDVDRate_LostFocus()
If Me.txtDVDRate.Text = "" Or Me.txtDVDRate.Text = "0" Then Me.txtDVDRate.Text = "0.00"
End Sub

Private Sub txtVCDRate_GotFocus()
If Me.txtVCDRate.Text = "" Or Me.txtVCDRate.Text = "0.00" Then Me.txtVCDRate.Text = ""
End Sub

Private Sub txtVCDRate_LostFocus()
If Me.txtVCDRate.Text = "" Or Me.txtVCDRate.Text = "0" Then Me.txtVCDRate.Text = "0.00"
End Sub
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
Call unlockF
Me.bttnEdit.Enabled = False
Me.bttnSave.Enabled = True
End Sub

Private Sub bttnSave_Click()
'
'    Dim ctr As Long
'    ctr = 1
'    Set rs = New ADODB.Recordset
'    With rs
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
'    !TransType = tType
'    !UserID = user
'    !TDate = Date
'    !TTime = Time
'    .Update
'
'    .Close
'End With


Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_RentalRate", cn, 1, 2
    
    If .EOF = True Then
        .AddNew
        !DVDAmount = Val(Me.txtDVDRate.Text)
        !VCDAmount = Val(Me.txtVCDRate.Text)
        .Update
    Else
        !DVDAmount = Me.txtDVDRate.Text
        !VCDAmount = Me.txtVCDRate.Text
        .Update
    End If
        MsgBox "Rental Rate Successfully Saved.", vbInformation + vbOKOnly
        .Close
        
        Call loadRecords
        Call lockF
        Me.bttnSave.Enabled = False
        Me.bttnEdit.Enabled = True
End With
End Sub
