VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO5492~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.4#0"; "CO3CF5~1.OCX"
Begin VB.Form frmThemes 
   Caption         =   "Themes"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdSave 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Picture         =   "frmThemes.frx":0000
   End
   Begin VB.OptionButton rbttn 
      Caption         =   "Office 2007"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton rbttn 
      Caption         =   "WinXP Luna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton rbttn 
      Caption         =   "WinXP Royale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.OptionButton rbttn 
      Caption         =   "Vista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin XtremeSuiteControls.PushButton cmdClose 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Picture         =   "frmThemes.frx":039A
   End
   Begin VB.Label Label1 
      Caption         =   "Theme Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   240
      Top             =   480
      _Version        =   851972
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim themes As String
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Sub Command2_Click()

End Sub

Private Sub cmdCLose_Click()

    Me.Hide
End Sub



Private Sub cmdSave_Click()

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Themes", cn, 1, 2
    If .EOF = True Then
        .AddNew
        !themes = themes
        .Update
    Else
        !themes = themes
        .Update
    End If
        .Close
        MsgBox "Themes Successfully Save.", vbInformation + vbOKOnly
End With
    Me.Hide
'*************************************
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub rbttn_Click(Index As Integer)

Select Case Index
    Case 0
        themes = "WinXP.Luna.cjstyles"
        
        MDIMain.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\WinXp.Luna.cjstyles", ""
        MDIMain.SkinFramework.ApplyWindow Me.hWnd
        MDIMain.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    Case 1
        themes = "WinXP.Royale.cjstyles"
        
        MDIMain.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\WinXP.Royale.cjstyles", ""
        MDIMain.SkinFramework.ApplyWindow Me.hWnd
        MDIMain.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    Case 2
        themes = "Vista.cjstyles"
        
        MDIMain.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\Vista.cjstyles", ""
        MDIMain.SkinFramework.ApplyWindow Me.hWnd
        MDIMain.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    Case 3
        themes = "Office2007.cjstyles"
        
        MDIMain.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\Office2007.cjstyles", ""
        MDIMain.SkinFramework.ApplyWindow Me.hWnd
        MDIMain.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End Select

End Sub


