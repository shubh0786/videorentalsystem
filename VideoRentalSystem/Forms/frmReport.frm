VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmReport 
   Caption         =   "Report Module"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboiCategory 
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin XtremeSuiteControls.PushButton cmdGenerate 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Generate"
      Appearance      =   6
   End
   Begin VB.Label Label1 
      Caption         =   "Select type first then click generate button to display the report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboiCategory_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdGenerate_Click()
If Combo_Empty(Me.cboiCategory) = True Then Exit Sub

Set rs = New ADODB.Recordset
rs.Open "Select * from qrequirement where category like '" & Me.cboiCategory.Text & "'", cn, 1, 2

End Sub

Private Sub Form_Load()

Me.cboiCategory.Text = "Available Discs"
Me.cboiCategory.Text = "Unreturn Discs"
Me.cboiCategory.Text = "Members"

End Sub
