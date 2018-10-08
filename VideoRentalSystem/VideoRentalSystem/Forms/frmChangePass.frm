VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.0#0"; "CODEJO~4.OCX"
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnOk 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
      _Version        =   983040
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "OK"
      Appearance      =   6
      Picture         =   "frmChangePass.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCancel 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
      _Version        =   983040
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancel"
      Appearance      =   6
      Picture         =   "frmChangePass.frx":039A
   End
   Begin XtremeSuiteControls.FlatEdit txtOld 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _Version        =   983040
      _ExtentX        =   4895
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
      Text            =   "FlatEdit1"
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNew 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
      _Version        =   983040
      _ExtentX        =   4895
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
      Text            =   "FlatEdit1"
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConfirm 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
      _Version        =   983040
      _ExtentX        =   4895
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
      Text            =   "FlatEdit1"
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
      _Version        =   983040
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Confirm Password:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
      _Version        =   983040
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "New Password:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   960
      Y2              =   960
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _Version        =   983040
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Old Password:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()

    If Text_Empty(Me.txtOld) = True Then Exit Sub
    If Text_Empty(Me.txtNew) = True Then Exit Sub
    If Text_Empty(Me.txtConfirm) = True Then Exit Sub
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users where username like '" & user & "'", cn, 1, 2
    
    If !Password <> Me.txtOld.Text Then
        MsgBox "Invalid Password", vbCritical + vbOKOnly
    Else
        If Me.txtNew.Text <> Me.txtConfirm.Text Then
            MsgBox "Password does not match.", vbExclamation + vbOKOnly
        Else
            !Password = Me.txtNew.Text
            .Update
            
            '***********Audit Trail
            wDone = "Change Password" & "/" & "User ID" & " - " & user
            Call auditT
            
            MsgBox "Password Changed!", vbInformation + vbOKOnly
            Unload Me
        End If
    End If
    .Close
End With
End Sub

Private Sub Form_Load()
With Me
    .txtConfirm.Text = ""
    .txtNew.Text = ""
    .txtOld.Text = ""
    
    .txtConfirm.PasswordChar = Chr(149)
    .txtNew.PasswordChar = Chr(149)
    .txtOld.PasswordChar = Chr(149)
End With
End Sub

