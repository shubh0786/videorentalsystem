VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO5492~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.4#0"; "CO3CF5~1.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Security"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "frmLogin.frx":628A
   ScaleHeight     =   2490
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _Version        =   851972
      _ExtentX        =   10186
      _ExtentY        =   4471
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      ItemCount       =   1
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "Label1"
      Item(0).Control(1)=   "txtPass"
      Item(0).Control(2)=   "Label2"
      Item(0).Control(3)=   "cmdOk"
      Item(0).Control(4)=   "cmdClose"
      Item(0).Control(5)=   "Label3"
      Item(0).Control(6)=   "lblAttempt"
      Item(0).Control(7)=   "lblCapsOnOff"
      Item(0).Control(8)=   "Label4"
      Item(0).Control(9)=   "txtUser"
      Begin XtremeSuiteControls.PushButton cmdOk 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
         _Version        =   851972
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&OK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmLogin.frx":6F54
      End
      Begin XtremeSuiteControls.PushButton cmdClose 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
         _Version        =   851972
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmLogin.frx":72EE
      End
      Begin XtremeSuiteControls.FlatEdit txtUser 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   3255
         _Version        =   851972
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin XtremeSuiteControls.FlatEdit txtPass 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   3255
         _Version        =   851972
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin XtremeSkinFramework.SkinFramework SkinFramework 
         Left            =   120
         Top             =   600
         _Version        =   851972
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Caps Lock:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Attempt:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblCapsOnOff 
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblAttempt 
         Caption         =   "3"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Password"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Username"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Sub cmdCLose_Click()
    If MsgBox("Exit application?", vbYesNo + vbQuestion) = vbYes Then End
End Sub
Private Sub cmdOk_Click()

    Call checkUser
    If Text_Empty(Me.txtUser) = True Then Exit Sub
    If Text_Empty(Me.txtPass) = True Then Exit Sub
    On Error GoTo err
    Static ctr As Integer
    Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Users where UserName like '" & txtUser.Text & "' and Password like '" & txtPass.Text & "' and Status like '" & "ACTIVE" & "'", cn, 1, 2
        
    If rs.EOF = True Then
        MsgBox "Access denied.", vbExclamation
        ctr = ctr + 1
        Me.lblAttempt.Caption = Val(Me.lblAttempt.Caption) - 1
    Else
        MsgBox "Access Granted."
                Set rs = New ADODB.Recordset
                        With rs
                            .Open "Select * from tbl_LogHistory", cn, 1, 2
       
                            .AddNew
                            !UserName = txtUser.Text
                            !LoginDate = Date
                            !loginTime = Time
                            !logOutTime = ""
                            !logoutDate = ""
                            !unitUsed = "Server"
                            
                            loginTime = Time
                            .Update
                End With
                    Set rs = Nothing
                        user = Me.txtUser.Text
                       MDIMain.Show
                        'frmSplash.Show
                Unload Me
   
                Exit Sub
    End If
        If ctr = 3 Then
            MsgBox "Good Bye."
            End
        End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
Private Sub Form_Activate()
    txtUser.SetFocus
    txtPass.PasswordChar = Chr(149)
End Sub
Private Sub Form_Load()
    TabControl1(0).Caption = "Login"
    
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Themes", cn, 1, 2

    SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\" + !themes, ""
    SkinFramework.ApplyWindow Me.hWnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
   End With
    
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
    If KeyAscii = 27 Then clear
'End If
End Sub
Sub clear()
    txtUsername.Text = ""
    txtPass.Text = ""
    txtUsername.SetFocus
End Sub

Sub checkUser()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users where UserName like '" & Me.txtUser.Text & "'", cn, 1, 2
    If .EOF = True Then
        OK = False
        Exit Sub
    Else
        If !UserType = "ADMINISTRATOR" Then OK = True
        If !UserType <> "ADMINISTRATOR" Then OK = False
    End If
    .Close
End With
End Sub

Private Sub DisplayCapsState()
    Dim CapsState As Long
        CapsState = GetKeyState(vbKeyCapital)
    If CapsState = 1 Then
        lblCapsOnOff.Caption = "On"
    Else
        lblCapsOnOff.Caption = "Off"
        End If
End Sub
Private Sub txtUser_Change()
        Call DisplayCapsState
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    cmdCLose_Click
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
    If KeyAscii = 27 Then clear
'End If
End Sub




