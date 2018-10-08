VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.4#0"; "COAAB5~1.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Video Rental System"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14070
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   0
      Picture         =   "frmSplash.frx":628A
      ScaleHeight     =   8595
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1080
         Top             =   840
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   360
         Top             =   600
      End
      Begin VideoRentalSystem.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   8160
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   873
         Value           =   0
         Theme           =   3
         TextStyle       =   2
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "U11D ProgressBar"
         TextEffectColor =   65280
         TextEffect      =   4
      End
      Begin XtremeSkinFramework.SkinFramework SkinFramework 
         Left            =   0
         Top             =   0
         _Version        =   851972
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Label lblconnection 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   8640
      Width           =   6735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()


Private Sub Form_Activate()
    Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Themes", cn, 1, 2

    frmLogin.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\" + !themes, ""
    frmLogin.SkinFramework.ApplyWindow Me.hWnd
    frmLogin.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
   End With
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Themes", cn, 1, 2

    frmLogin.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\" + !themes, ""
    frmLogin.SkinFramework.ApplyWindow Me.hWnd
    frmLogin.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
   End With
    
End Sub

Private Sub Form_Resize()
ProgressBar1.Top = ProgressBar1.Top + 7
ProgressBar1.Width = ProgressBar1.Width + 69
lblconnection.Top = lblconnection.Top + 7
lblconnection.Left = lblconnection.Left + 26
End Sub



Private Sub Timer1_Timer()
Dim n As String
Static count As Integer
    With ProgressBar1
    .value = count
            ProgressBar1.value = ProgressBar1.value + 1
        If ProgressBar1.value = 3 Then
            lblconnection.Caption = "Initializing Application . . ."
        ElseIf ProgressBar1.value = 15 Then
            lblconnection.Caption = "Please wait . . ."
        ElseIf ProgressBar1.value = 25 Then
            lblconnection.Caption = "Checking Database . . ."
        ElseIf ProgressBar1.value = 30 Then
            lblconnection.Caption = "Checking Database . . ."
        ElseIf ProgressBar1.value = 40 Then
            lblconnection.Caption = "Connecting Database . . ."
        ElseIf ProgressBar1.value = 50 Then
            lblconnection.Caption = "Connecting Database . . ."
        ElseIf ProgressBar1.value = 75 Then
            lblconnection.Caption = "Connecting Database . . ."
        ElseIf ProgressBar1.value = 85 Then
            lblconnection.Caption = "Loading form . . ."
        ElseIf ProgressBar1.value = 93 Then
            lblconnection.Caption = "Finishing . . ."
        ElseIf ProgressBar1.value = 100 Then
            MDIMain.Show
            Unload Me
End If
End With
count = count + 1
        If count >= Timer1.Interval Then
            Timer1.Enabled = False
            Me.Hide
            End If
End Sub

