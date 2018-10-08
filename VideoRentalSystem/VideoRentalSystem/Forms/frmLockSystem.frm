VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Object = "{DC808AFB-A306-4615-8D4B-83064675D70D}#2.0#0"; "SideTab.ocx"
Begin VB.Form frmLockSystem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmLockSystem.frx":0000
   ScaleHeight     =   10800
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   7680
   End
   Begin Project4.Joemel_SideTab tab1 
      Height          =   1095
      Left            =   6120
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1931
      Caption         =   "Enter User Password"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Arial"
      FontSize        =   9.75
      ForeColor       =   16777215
      BackColor       =   12632256
      Begin XtremeSuiteControls.FlatEdit txtPass 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton bttnUnlock 
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _Version        =   851972
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Unlock"
         BackColor       =   -2147483633
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
         Picture         =   "frmLockSystem.frx":2207D2
      End
   End
End
Attribute VB_Name = "frmLockSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnUnlock_Click()


 If Text_Empty(Me.txtPass) = True Then Exit Sub

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users where Username like '" & user & "' and Password like '" & txtPass.Text & "'", cn, 1, 2
    
        If .EOF Then
            If MsgBox("Access Denied.", vbExclamation + vbOKOnly) = vbOK Then txtPass.Text = ""
        Else
            MsgBox "Successfuly Unlock.", vbInformation + vbOKOnly
                frmLockSystem.Hide
                MDIMain.Show
        End If
        .Close
End With
End Sub

Private Sub Form_Load()
Me.WindowState = 2
txtPass.PasswordChar = Chr(149)
tab1.Visible = True
Timer1.Interval = 1000
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tab1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Picture1_Click()
On Error Resume Next
Timer1.Enabled = True
    tab1.Visible = True
End Sub

Private Sub Timer1_Timer()
Dim X As Integer
    If X = 60 Then
        Timer1.Enabled = False
        tab1.Visible = False
    Else
        X = X + 1
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnUnlock_Click

'End If
End Sub

