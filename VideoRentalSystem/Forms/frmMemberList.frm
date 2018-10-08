VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmMemberList 
   Caption         =   "Member List"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   12030
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
         Height          =   4095
         Left            =   480
         TabIndex        =   1
         Top             =   1680
         Width           =   11640
         _Version        =   851972
         _ExtentX        =   20532
         _ExtentY        =   7223
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
      Begin XtremeSuiteControls.PushButton bttnViewRecord 
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   6000
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
         Picture         =   "frmMemberList.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   10815
         _Version        =   851972
         _ExtentX        =   19076
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Search or Filter"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cboFilter 
            Height          =   360
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
            _ExtentY        =   635
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
         End
         Begin XtremeSuiteControls.PushButton bttnRefresh 
            Height          =   375
            Left            =   7920
            TabIndex        =   5
            Top             =   240
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
            Picture         =   "frmMemberList.frx":039A
         End
         Begin XtremeSuiteControls.FlatEdit txtSearch 
            Height          =   375
            Left            =   3360
            TabIndex        =   6
            Top             =   240
            Width           =   4335
            _Version        =   851972
            _ExtentX        =   7646
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
         End
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   4200
         Picture         =   "frmMemberList.frx":0B14
         Top             =   5880
         Width           =   8730
      End
   End
End
Attribute VB_Name = "frmMemberList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Me.lvView.Width = MDIMain.Width
    Me.GroupBox2.Width = MDIMain.Width
    Me.TabControl1.Width = MDIMain.Width

    Call loadRecords
    Call getRecords

End Sub


Sub getRecords()
    Me.lvView.ListItems.clear
    Set rs = New ADODB.Recordset
With rs
   
    If Me.cboFilter.Text = "ID Number" Then
        .Open "Select * from tbl_Members where memberID like '" & Me.txtSearch.Text & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
    ElseIf Me.cboFilter.Text = "Last Name" Then
        .Open "Select * from tbl_Members where Lname like '" & Me.txtSearch.Text & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
    End If

    Do Until .EOF
    Set lst = lvView.ListItems.Add(, , !MemberID)
        lst.ListSubItems.Add , , !Lname & ", " & !Fname & " " & !Mname
        lst.ListSubItems.Add , , !Bdate
        lst.ListSubItems.Add , , !Sex
        lst.ListSubItems.Add , , !Age
        lst.ListSubItems.Add , , !CivilStatus
        lst.ListSubItems.Add , , !CurAdd
        lst.ListSubItems.Add , , !PerAdd
        lst.ListSubItems.Add , , !Cellphone
        lst.ListSubItems.Add , , !Email
        lst.ListSubItems.Add , , !IDPresented
        lst.ListSubItems.Add , , !IDCardNo
        lst.ListSubItems.Add , , !OtherID
        lst.ListSubItems.Add , , !ExLname & ", " & !ExFname & " " & !ExMname
        lst.ListSubItems.Add , , !Relation
        lst.ListSubItems.Add , , !ExCurAdd
        lst.ListSubItems.Add , , !ExCellPhone
        lst.ListSubItems.Add , , !ExEmail
        
        .MoveNext
    Loop
End With
End Sub

Private Sub txtSearch_Change()
    Call getRecords
End Sub

Private Sub bttnRefresh_Click()
    Call getRecords
End Sub


Sub viewRecords()
On Error GoTo err
    frmMemberInfo.lblID.Caption = Me.lvView.SelectedItem
    frmMemberInfo.lblName.Caption = Me.lvView.SelectedItem.ListSubItems(1)
    frmMemberInfo.lblBirth.Caption = Me.lvView.SelectedItem.ListSubItems(2)
    frmMemberInfo.lblGender.Caption = Me.lvView.SelectedItem.ListSubItems(3)
    frmMemberInfo.lblAge.Caption = Me.lvView.SelectedItem.ListSubItems(4)
    frmMemberInfo.lblCivil.Caption = Me.lvView.SelectedItem.ListSubItems(5)
    frmMemberInfo.lblAddress.Caption = Me.lvView.SelectedItem.ListSubItems(6)
    frmMemberInfo.lblCelNo.Caption = Me.lvView.SelectedItem.ListSubItems(8)
    frmMemberInfo.lblEMail.Caption = Me.lvView.SelectedItem.ListSubItems(9)
    frmMemberInfo.lblIDPres.Caption = Me.lvView.SelectedItem.ListSubItems(10)
    frmMemberInfo.lblCardNo.Caption = Me.lvView.SelectedItem.ListSubItems(11)
    frmMemberInfo.lblExName.Caption = Me.lvView.SelectedItem.ListSubItems(13)
    frmMemberInfo.lblExRelation.Caption = Me.lvView.SelectedItem.ListSubItems(14)
    frmMemberInfo.lblExCurAddress.Caption = Me.lvView.SelectedItem.ListSubItems(15)
    
       Set rs = New ADODB.Recordset
    With rs
   
    If Me.cboFilter.Text = "ID Number" Then
        .Open "Select * from tbl_Members where memberID like '" & Me.lvView.SelectedItem & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
         Set frmMemberInfo.Image2.DataSource = rs
            frmMemberInfo.Image2.DataField = "pic"
            .Close
    ElseIf Me.cboFilter.Text = "Last Name" Then
        .Open "Select * from tbl_Members where Lname like '" & Me.lvView.SelectedItem & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
         Set frmMemberInfo.Image2.DataSource = rs
                frmMemberInfo.Image2.DataField = "pic"
                .Close
    End If
               
        End With
        
        frmMemberInfo.Image1.Height = 2295
        frmMemberInfo.Image1.Width = 3735
        frmMemberInfo.Image1.Stretch = True
        frmMemberInfo.Show vbModal
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
.Open "Select * from tbl_Members where status like '" & "ACTIVE" & "'", cn, 1, 2
    With Me.lvView
    .ColumnHeaders.clear
    .ListItems.clear
    .ColumnHeaders.Add , , "Member ID", 1000
    .ColumnHeaders.Add , , "Name", 1500
    .ColumnHeaders.Add , , "Birth Date", 1000
    .ColumnHeaders.Add , , "Gender", 1500
    .ColumnHeaders.Add , , "Age", 1000
    .ColumnHeaders.Add , , "Civil Status", 1000
    .ColumnHeaders.Add , , "Current Address", 1500
    .ColumnHeaders.Add , , "Permanent Address", 1500
    .ColumnHeaders.Add , , "CellPhone No.", 1500
    .ColumnHeaders.Add , , "Email Address", 1500
    .ColumnHeaders.Add , , "ID Presented", 1500
    .ColumnHeaders.Add , , "ID CardNo", 1500
    .ColumnHeaders.Add , , "Other ID", 1500
    .ColumnHeaders.Add , , "Extension Name", 1500
    .ColumnHeaders.Add , , "Relation", 1500
    .ColumnHeaders.Add , , "Extension Current Address", 1500
    .ColumnHeaders.Add , , "Extension CellPhone No", 1500
    .ColumnHeaders.Add , , "Extension Email Address", 1500
End With

    With Me.cboFilter
        .AddItem "ID Number"
        .AddItem "Last Name"
        .Text = "ID Number"
    End With

     Do Until .EOF
    Set lst = lvView.ListItems.Add(, , !MemberID)
        lst.ListSubItems.Add , , !Lname & ", " & !Fname & " " & !Mname
        lst.ListSubItems.Add , , !Bdate
        lst.ListSubItems.Add , , !Sex
        lst.ListSubItems.Add , , !Age
        lst.ListSubItems.Add , , !CivilStatus
        lst.ListSubItems.Add , , !CurAdd
        lst.ListSubItems.Add , , !PerAdd
        lst.ListSubItems.Add , , !Cellphone
        lst.ListSubItems.Add , , !Email
        lst.ListSubItems.Add , , !IDPresented
        lst.ListSubItems.Add , , !IDCardNo
        lst.ListSubItems.Add , , !OtherID
        lst.ListSubItems.Add , , !ExLname & ", " & !ExFname & " " & !ExMname
        lst.ListSubItems.Add , , !Relation
        lst.ListSubItems.Add , , !ExCurAdd
        lst.ListSubItems.Add , , !ExCellPhone
        lst.ListSubItems.Add , , !ExEmail
        
        .MoveNext
    Loop
    
    End With
End Sub
