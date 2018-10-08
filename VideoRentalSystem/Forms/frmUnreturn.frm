VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmUnreturn 
   Caption         =   "Unreturn"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   10860
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   9855
      _Version        =   851972
      _ExtentX        =   17383
      _ExtentY        =   5741
      _StockProps     =   77
      BackColor       =   -2147483643
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox groupbx 
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   6135
      _Version        =   851972
      _ExtentX        =   10821
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtSearch 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboFilter 
         Height          =   360
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "Title"
      End
      Begin VB.Label Label2 
         Caption         =   "Search:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.PushButton bttnRefresh 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   840
      Width           =   2175
      _Version        =   851972
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Refresh"
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
      Picture         =   "frmUnreturn.frx":0000
   End
End
Attribute VB_Name = "frmUnreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me.cboFilter
    .AddItem "Title"
    .AddItem "CD/DVD ID"
    
    End With
    Me.ListView1.Width = MDIMain.Width
With Me.ListView1
    .ColumnHeaders.Add , , "CD Number", 1500
    .ColumnHeaders.Add , , "CD ID", 1500
    .ColumnHeaders.Add , , "Rent Date", 2000
    .ColumnHeaders.Add , , "Return Date", 1500
    .ColumnHeaders.Add , , "Title", 1500
    .ColumnHeaders.Add , , "Genre", 1500
    .ColumnHeaders.Add , , "Director", 1500
    .ColumnHeaders.Add , , "Actor/Actress", 1500
    .ColumnHeaders.Add , , "CD Type", 1500
    .ColumnHeaders.Add , , "Category", 1500
    .ColumnHeaders.Add , , "Borrower ID", 1500
    .ColumnHeaders.Add , , "Borrower Name", 1500
End With

Call loadData
End Sub


Sub loadData()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from QUnreturn", cn, 1, 2
    Do Until .EOF
        Set lst = Me.ListView1.ListItems.Add(, , !Ino)
            'With lst
                lst.ListSubItems.Add , , !CD_ID
                lst.ListSubItems.Add , , !rentDate
                lst.ListSubItems.Add , , !ReturnDate
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !Ggenre
                lst.ListSubItems.Add , , !director
                lst.ListSubItems.Add , , !actor
                lst.ListSubItems.Add , , !CDType
                lst.ListSubItems.Add , , !Category
                lst.ListSubItems.Add , , !MemberID
                lst.ListSubItems.Add , , !Lname & "," & !Fname & "" & !Mname
            'End With
        .MoveNext
    Loop
End With
End Sub
Private Sub txtSearch_Change()
    Me.ListView1.ListItems.clear
    Set rs = New ADODB.Recordset
    With rs
    
        If Me.cboFilter.Text = "Title" Then .Open "Select * from QUnreturn where Title like '" & Me.txtSearch.Text & "%'", cn, 1, 2
        If Me.cboFilter.Text = "CD/DVD ID" Then .Open "Select * from QUnreturn where INo like '" & Me.txtSearch.Text & "%'", cn, 1, 2
         Do Until .EOF
        Set lst = Me.ListView1.ListItems.Add(, , !Ino)
            'With lst
                lst.ListSubItems.Add , , !CD_ID
                lst.ListSubItems.Add , , !rentDate
                lst.ListSubItems.Add , , !ReturnDate
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !Ggenre
                lst.ListSubItems.Add , , !director
                lst.ListSubItems.Add , , !actor
                lst.ListSubItems.Add , , !CDType
                lst.ListSubItems.Add , , !Category
                lst.ListSubItems.Add , , !MemberID
                lst.ListSubItems.Add , , !Lname & "," & !Fname & "" & !Mname
            'End With
        .MoveNext
    Loop
            .Close
    End With
End Sub


Private Sub cboFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 0 Then KeyAscii = 0
End Sub
