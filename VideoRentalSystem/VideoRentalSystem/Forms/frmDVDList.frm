VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmDVDList 
   Caption         =   "CD/DVD List"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   13125
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _Version        =   851972
      _ExtentX        =   23098
      _ExtentY        =   14420
      _StockProps     =   68
      Appearance      =   3
      Color           =   2
      Begin XtremeSuiteControls.ListView lvView 
         Height          =   4695
         Left            =   480
         TabIndex        =   3
         Top             =   1680
         Width           =   11640
         _Version        =   851972
         _ExtentX        =   20532
         _ExtentY        =   8281
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton bttnRefresh 
         Height          =   375
         Left            =   6840
         TabIndex        =   7
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
         Picture         =   "frmDVDList.frx":0000
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
      Begin XtremeSuiteControls.PushButton bttnViewRecord 
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   6480
         Width           =   2655
         _Version        =   851972
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "View Record"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmDVDList.frx":077A
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   4200
         Picture         =   "frmDVDList.frx":0B14
         Top             =   6480
         Width           =   9225
      End
   End
End
Attribute VB_Name = "frmDVDList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call loadRecords
    Me.lvView.Width = MDIMain.Width
    Me.groupbx.Width = MDIMain.Width
    Me.TabControl1.Width = MDIMain.Width
End Sub

Sub viewRecords()
On Error GoTo err
    frmDVDInfo.lblID.Caption = Me.lvView.SelectedItem
    frmDVDInfo.lblTitle.Caption = Me.lvView.SelectedItem.ListSubItems(2)
    frmDVDInfo.lblGenre.Caption = Me.lvView.SelectedItem.ListSubItems(3)
    frmDVDInfo.lblActor.Caption = Me.lvView.SelectedItem.ListSubItems(5)
    frmDVDInfo.lblDirector.Caption = Me.lvView.SelectedItem.ListSubItems(4)
    frmDVDInfo.lblYrRelease.Caption = Me.lvView.SelectedItem.ListSubItems(6)
    frmDVDInfo.lblCdType.Caption = Me.lvView.SelectedItem.ListSubItems(7)
    frmDVDInfo.lblCategory.Caption = Me.lvView.SelectedItem.ListSubItems(8)
    frmDVDInfo.lblNoCopy.Caption = Me.lvView.SelectedItem.ListSubItems(10)

    Set rs = New ADODB.Recordset
    With rs
    Dim ctr As Integer
    ctr = 0
        .Open "Select * from tbl_IndividualCDList where CD_ID like '" & Me.lvView.SelectedItem & "' and status like 'AVAILABLE'", cn, 1, 2
    
    Do Until .EOF
        ctr = ctr + 1
        .MoveNext
    Loop
        .Close
    End With
        frmDVDInfo.lblAvaiCopy.Caption = ctr
        
        Set rs = New ADODB.Recordset
        With rs
            .Open "Select * from tbl_CDList where CD_ID like '" & Me.lvView.SelectedItem & "'", cn, 1, 2
    
                Set frmDVDInfo.Image1.DataSource = rs
                frmDVDInfo.Image1.DataField = "pic"
                .Close
        End With
        
        frmDVDInfo.Image1.Height = 2295
        frmDVDInfo.Image1.Width = 3735
        frmDVDInfo.Image1.Stretch = True
        frmDVDInfo.Show vbModal
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub bttnViewRecord_Click()
    Call viewRecords
End Sub

Private Sub cboFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 0 Then KeyAscii = 0
End Sub


Private Sub lvView_DblClick()
On Error GoTo err
    If Me.lvView.SelectedItem <> "" Then
        Call viewRecords
    Else
    End If
    Exit Sub
err:
    MsgBox "No item selected.", vbInformation + vbOKOnly
End Sub

Private Sub lvView_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Me.bttnViewRecord.Enabled = True
End Sub

Sub loadRecords()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from QCDList", cn, 1, 1
    With Me.lvView
    .ColumnHeaders.clear
    .ListItems.clear
    .ColumnHeaders.Add , , "CD ID", 1300
    .ColumnHeaders.Add , , "Quantity", 1300
    .ColumnHeaders.Add , , "Title", 4300
    .ColumnHeaders.Add , , "Genre", 1300
    .ColumnHeaders.Add , , "Director", 1300
    .ColumnHeaders.Add , , "Actor/Actress", 1300
    .ColumnHeaders.Add , , "Year Realese", 1300
    .ColumnHeaders.Add , , "CD Type", 1300
    .ColumnHeaders.Add , , "Rent Duration", 1300
    .ColumnHeaders.Add , , "Category", 1300
    .ColumnHeaders.Add , , "", 1
End With

With Me.cboFilter
    .AddItem "Title"
    .AddItem "CD/DVD ID"
End With
    Do Until .EOF
            Set lst = Me.lvView.ListItems.Add(, , !CD_ID)
            lst.ListSubItems.Add , , !Qnty
            lst.ListSubItems.Add , , !Title
            lst.ListSubItems.Add , , !Ggenre
            lst.ListSubItems.Add , , !director
            lst.ListSubItems.Add , , !actor
            lst.ListSubItems.Add , , !YearRelease
            lst.ListSubItems.Add , , !CDType
            lst.ListSubItems.Add , , !goodFor
            lst.ListSubItems.Add , , !Category
            lst.ListSubItems.Add , , !Qnty
        .MoveNext
    Loop
    .Close
End With
End Sub



Private Sub pushEdit_Click()

End Sub

Private Sub txtSearch_Change()
    Me.lvView.ListItems.clear
    Set rs = New ADODB.Recordset
    With rs
    
        If Me.cboFilter.Text = "Title" Then .Open "Select * from QCDList where title like '" & Me.txtSearch.Text & "%'", cn, 1, 2
        If Me.cboFilter.Text = "CD/DVD ID" Then .Open "Select * from QCDlist where CD_ID like '" & Me.txtSearch.Text & "%'", cn, 1, 2
        Do Until .EOF
            Set lst = Me.lvView.ListItems.Add(, , !CD_ID)
            lst.ListSubItems.Add , , !Qnty
            lst.ListSubItems.Add , , !Title
            lst.ListSubItems.Add , , !Ggenre
            lst.ListSubItems.Add , , !director
            lst.ListSubItems.Add , , !actor
            lst.ListSubItems.Add , , !YearRelease
            lst.ListSubItems.Add , , !CDType
            lst.ListSubItems.Add , , !goodFor
            lst.ListSubItems.Add , , !Category
            lst.ListSubItems.Add , , !Qnty
            .MoveNext
            Loop
            .Close
    End With
End Sub


