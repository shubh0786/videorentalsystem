VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CODEJO~1.OCX"
Begin VB.Form frmMembers 
   Caption         =   "Modify Members"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      _Version        =   851972
      _ExtentX        =   21193
      _ExtentY        =   11880
      _StockProps     =   79
      Caption         =   "List of Members"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Begin XtremeSuiteControls.ListView lvView 
         Height          =   4095
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   10815
         _Version        =   851972
         _ExtentX        =   19076
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   4680
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
            Left            =   480
            TabIndex        =   3
            Top             =   240
            Width           =   2655
            _Version        =   851972
            _ExtentX        =   4683
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
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            _Version        =   851972
            _ExtentX        =   2355
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
            TextImageRelation=   1
         End
         Begin XtremeSuiteControls.FlatEdit txtSearch 
            Height          =   375
            Left            =   3360
            TabIndex        =   5
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
      Begin XtremeSuiteControls.PushButton bttnNew 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   6000
         Width           =   2655
         _Version        =   851972
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "New"
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
         Picture         =   "frmMembers.frx":0000
      End
      Begin XtremeSuiteControls.PushButton bttnEdit 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   6000
         Width           =   2655
         _Version        =   851972
         _ExtentX        =   4683
         _ExtentY        =   661
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
         Picture         =   "frmMembers.frx":039A
      End
      Begin XtremeSuiteControls.PushButton bttnDel 
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   6000
         Width           =   2415
         _Version        =   851972
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Delete"
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
         Picture         =   "frmMembers.frx":0734
      End
      Begin XtremeSuiteControls.PushButton bttnCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   6000
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cancel"
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
         Picture         =   "frmMembers.frx":0ACE
      End
      Begin VB.Label Label3 
         Caption         =   "Please double Click the items above."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   10
         Top             =   5520
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnt As Integer

Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub bttnCancel_Click()
    If Me.bttnCancel.Caption = "Close" Then Unload Me
    
End Sub
Private Sub bttnDel_Click()
'Me.lvView.Enabled = True
On Error Resume Next

' Set rs = New ADODB.Recordset
'
''    If Me.cboFilter.Text = "ID Number" Then
'        rs.Open "Select * from tbl_Members where memberID like '" & Me.txtSearch.Text & "%'", cn, 1, 2
'        rs.Close
'
If text_empty(Me.txtSearch) = True Then

'Exit Sub
ElseIf lvView.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation
'Exit Sub

Else
    varx = Me.lvView.SelectedItem.Index
    If MsgBox("Delete Records?", vbQuestion + vbYesNo) = vbYes Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Members where MemberID like '" & Me.lvView.ListItems(varx).Text & "'", cn, 1, 2
         .Open "Select * from tbl_Members where MemberID like '" & Me.txtSearch.Text & "'", cn, 1, 2
            !Status = "DELETED"
            .Update
                MsgBox "Record has been deleted.", vbInformation + vbOKOnly
        .Close
      wDone = "Deleted a Member" & "/" & "Member ID" & " - " & Me.lvView.ListItems(varx).Text
            Call auditT
    End With
    Me.lvView.ListItems.Remove (varx)
    End If
    End If
End Sub


Private Sub bttnEdit_Click()

If text_empty(Me.txtSearch) = True Then Exit Sub

Me.lvView.Enabled = True

If lvView.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation
Label3.Visible = True
Exit Sub


''    Dim search As String
''    search = InputBox("Enter Member ID:")
''    Set rs = New ADODB.Recordset
''    rs.Open "Select * from tbl_Members where memberID like '" & search & "' and status like '" & "ACTIVE" & "'", cn, 1, 2
''    If rs.EOF = True Then
''        MsgBox "No record found!", vbExclamation
''    Else
''        frmNewMember.txtMemberID.Text = rs!MemberID
''
''    End If

'On Error GoTo err
'
'    Set rs = New ADODB.Recordset
'    rs.Open "Select * from tbl_Members where memberID like '" & Me.txtSearch.Text & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
'
'    rs!Age = frmNewMember.txtAge.Text
'    rs!Sex = frmNewMember.cboGender.Text
'    rs.Update
'    MsgBox "Record UPDATED!", vbInformation
'   loadRecords ""
'    Exit Sub
'
'err:
'    MsgBox err.Description, vbCritical
End Sub

Private Sub bttnNew_Click()
    frmNewMember.Show vbModal
    
End Sub

Private Sub bttnRefresh_Click()
   Call getRecords
End Sub

Private Sub cboFilter_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.lvView.Width = MDIMain.Width
    Me.GroupBox2.Width = MDIMain.Width
    Me.GroupBox1.Width = MDIMain.Width
    Me.Height = 7290
    Label3.Visible = False
  Call loadRecords
''   Call getRecords
   
   Me.lvView.Enabled = False
 Me.bttnCancel.Caption = "Close"
''
''    With Me.cboFilter
''        '.AddItem "ID Number"
''       ' .AddItem "Last Name"
''        '.Text = "ID Number"
''    End With

'   bttnEdit.Enabled = True
'     bttnNew.Enabled = True
'       bttnDel.Enabled = True
       
    
End Sub



Sub loadRecords()
    Set rs = New ADODB.Recordset
With rs
.Open "Select * from tbl_Members where status like '" & "ACTIVE" & "'", cn, 1, 2
    With lvView
    .ColumnHeaders.clear
    .ListItems.clear
    .ColumnHeaders.Add , , "Member ID", 1000
    .ColumnHeaders.Add , , "Last Name", 1000
    .ColumnHeaders.Add , , "First Name", 1000
    .ColumnHeaders.Add , , "Middle Name", 1000
    .ColumnHeaders.Add , , "Birth Date", 1500
    .ColumnHeaders.Add , , "Gender", 1500
    .ColumnHeaders.Add , , "Age", 1500
    .ColumnHeaders.Add , , "Civil Status", 1500
    .ColumnHeaders.Add , , "Current Address", 1500
    .ColumnHeaders.Add , , "Current ZIP Code", 1000
    .ColumnHeaders.Add , , "Permanent Address", 1500
    .ColumnHeaders.Add , , "Permanent ZIP Code", 1000
    .ColumnHeaders.Add , , "CellPhone No.", 1500
    .ColumnHeaders.Add , , "Email Address", 1500
    .ColumnHeaders.Add , , "ID Presented", 1500
    .ColumnHeaders.Add , , "ID CardNo", 1500
    .ColumnHeaders.Add , , "Other ID", 1500
    .ColumnHeaders.Add , , "Extension Last Name", 1000
    .ColumnHeaders.Add , , "Extension First Name", 1000
    .ColumnHeaders.Add , , "Extension Middle Name", 1000
    .ColumnHeaders.Add , , "Relation", 1500
    .ColumnHeaders.Add , , "Extension Current Address", 1500
     .ColumnHeaders.Add , , "Extension ZIP Code", 1000
    .ColumnHeaders.Add , , "Extension CellPhone No", 1500
    .ColumnHeaders.Add , , "Extension Email Address", 1500
    .ColumnHeaders.Add , , "Date Applied", 1500
End With

    With Me.cboFilter
        .AddItem "ID Number"
        .AddItem "Last Name"
    End With

     Do Until .EOF
     Set lst = lvView.ListItems.Add(, , !MemberID)
        lst.ListSubItems.Add , , !Lname
        lst.ListSubItems.Add , , !Fname
        lst.ListSubItems.Add , , !Mname
        lst.ListSubItems.Add , , !Bdate
        lst.ListSubItems.Add , , !Sex
        lst.ListSubItems.Add , , !Age
        lst.ListSubItems.Add , , !CivilStatus
        lst.ListSubItems.Add , , !CurAdd
         lst.ListSubItems.Add , , !CurZIP
        lst.ListSubItems.Add , , !PerAdd
        lst.ListSubItems.Add , , !PerZIP
        lst.ListSubItems.Add , , !Cellphone
        lst.ListSubItems.Add , , !Email
        lst.ListSubItems.Add , , !IDPresented
        lst.ListSubItems.Add , , !IDCardNo
        lst.ListSubItems.Add , , !OtherID
        lst.ListSubItems.Add , , !ExLname
        lst.ListSubItems.Add , , !ExFname
        lst.ListSubItems.Add , , !ExMname
        lst.ListSubItems.Add , , !Relation
        lst.ListSubItems.Add , , !ExCurAdd
        lst.ListSubItems.Add , , !ExZIP
        lst.ListSubItems.Add , , !ExCellPhone
        lst.ListSubItems.Add , , !ExEmail
        lst.ListSubItems.Add , , !DateApplied
        .MoveNext
    Loop
    
    End With
End Sub


Private Sub lvView_DblClick()
    Call viewRecords
End Sub


Private Sub txtSearch_Change()
   Call getRecords
End Sub
Sub viewRecords()
'On Error GoTo err

 frmModifyMember.dtpApplication.Enabled = False

    frmModifyMember.txtMemberID.Text = Me.lvView.SelectedItem
     frmModifyMember.txtLName.Text = Me.lvView.SelectedItem.ListSubItems(1)
    frmModifyMember.txtFName.Text = Me.lvView.SelectedItem.ListSubItems(2)
     frmModifyMember.txtMName.Text = Me.lvView.SelectedItem.ListSubItems(3)
     frmModifyMember.dtBdate.Value = Me.lvView.SelectedItem.ListSubItems(4)
      frmModifyMember.cboGender.Text = Me.lvView.SelectedItem.ListSubItems(5)
     frmModifyMember.txtAge.Text = Me.lvView.SelectedItem.ListSubItems(6)
    
   If frmModifyMember.rBttnSingle.Caption = Me.lvView.SelectedItem.ListSubItems(7) Then
         frmModifyMember.rBttnSingle.Value = True
    ElseIf frmModifyMember.rBttnMarried.Caption = Me.lvView.SelectedItem.ListSubItems(7) Then
          frmModifyMember.rBttnMarried.Value = True
      ElseIf frmModifyMember.rBttnWidow.Caption = Me.lvView.SelectedItem.ListSubItems(7) Then
          frmModifyMember.rBttnWidow.Value = True
    End If
   
       frmModifyMember.txtCurAdd.Text = Me.lvView.SelectedItem.ListSubItems(8)
        frmModifyMember.txtCurZip.Text = Me.lvView.SelectedItem.ListSubItems(9)
     frmModifyMember.txtPerAdd.Text = Me.lvView.SelectedItem.ListSubItems(10)
 
     frmModifyMember.txtPerZip.Text = Me.lvView.SelectedItem.ListSubItems(11)
     
     frmModifyMember.txtCelNo.Text = Me.lvView.SelectedItem.ListSubItems(12)
     frmModifyMember.txtEmailAdd.Text = Me.lvView.SelectedItem.ListSubItems(13)
          frmModifyMember.cboID.Text = Me.lvView.SelectedItem.ListSubItems(14)
      frmModifyMember.txtIDNo = Me.lvView.SelectedItem.ListSubItems(15)
      
       
    If frmModifyMember.rBttn1.Caption = Me.lvView.SelectedItem.ListSubItems(16) Then
         frmModifyMember.rBttn1.Value = True
    ElseIf frmModifyMember.rBttn2.Caption = Me.lvView.SelectedItem.ListSubItems(16) Then
          frmModifyMember.rBttn2.Value = True
      ElseIf frmModifyMember.rBttn3.Caption = Me.lvView.SelectedItem.ListSubItems(16) Then
          frmModifyMember.rBttn3.Value = True
    End If
  
     frmModifyMember.txtExLName.Text = Me.lvView.SelectedItem.ListSubItems(17)
     frmModifyMember.txtExFName.Text = Me.lvView.SelectedItem.ListSubItems(18)
     frmModifyMember.txtExMName.Text = Me.lvView.SelectedItem.ListSubItems(19)
       frmModifyMember.txtRelation.Text = Me.lvView.SelectedItem.ListSubItems(20)
        frmModifyMember.txtExCurAdd.Text = Me.lvView.SelectedItem.ListSubItems(21)

     frmModifyMember.txtExCurZip.Text = Me.lvView.SelectedItem.ListSubItems(22)
    frmModifyMember.txtExCelNo.Text = Me.lvView.SelectedItem.ListSubItems(23)
     frmModifyMember.txtExEmailAdd.Text = Me.lvView.SelectedItem.ListSubItems(24)

        frmModifyMember.dtpApplication.Value = Me.lvView.SelectedItem.ListSubItems(25)
      
    
 
        Set rs = New ADODB.Recordset
        With rs
          .Open "Select * from tbl_Members where memberID like '" & Me.lvView.SelectedItem & "%' and status like '" & "ACTIVE" & "'", cn, 1, 2
    
                Set frmModifyMember.Image1.DataSource = rs
                 frmModifyMember.Image1.DataField = "pic"
                .Close
        End With
        
         frmModifyMember.Image1.Stretch = True
         frmModifyMember.Show vbModal
    Exit Sub
'err:
'    MsgBox err.Description
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
        lst.ListSubItems.Add , , !Lname
        lst.ListSubItems.Add , , !Fname
        lst.ListSubItems.Add , , !Mname
        lst.ListSubItems.Add , , !Bdate
        lst.ListSubItems.Add , , !Sex
        lst.ListSubItems.Add , , !Age
        lst.ListSubItems.Add , , !CivilStatus
        lst.ListSubItems.Add , , !CurAdd
         lst.ListSubItems.Add , , !CurZIP
        lst.ListSubItems.Add , , !PerAdd
        lst.ListSubItems.Add , , !PerZIP
        lst.ListSubItems.Add , , !Cellphone
        lst.ListSubItems.Add , , !Email
        lst.ListSubItems.Add , , !IDPresented
        lst.ListSubItems.Add , , !IDCardNo
        lst.ListSubItems.Add , , !OtherID
        lst.ListSubItems.Add , , !ExLname
        lst.ListSubItems.Add , , !ExFname
        lst.ListSubItems.Add , , !ExMname
        lst.ListSubItems.Add , , !Relation
        lst.ListSubItems.Add , , !ExCurAdd
        lst.ListSubItems.Add , , !ExZIP
        lst.ListSubItems.Add , , !ExCellPhone
        lst.ListSubItems.Add , , !ExEmail
        lst.ListSubItems.Add , , !DateApplied
        .MoveNext
    Loop
End With
End Sub

