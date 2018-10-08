VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.0#0"; "CODEJO~4.OCX"
Begin VB.Form frmArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archive"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   11520
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _Version        =   983040
      _ExtentX        =   20346
      _ExtentY        =   11456
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.FixedTabWidth=   248
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Members"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "ListView1"
      Item(0).Control(1)=   "btnRestoreMember"
      Item(1).Caption =   "DISC"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "ListView3"
      Item(1).Control(1)=   "btnRestoreD"
      Item(2).Caption =   "User"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "ListView2"
      Item(2).Control(1)=   "btnRestoreUser"
      Begin XtremeSuiteControls.ListView ListView1 
         Height          =   4815
         Left            =   -69880
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   11295
         _Version        =   983040
         _ExtentX        =   19923
         _ExtentY        =   8493
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListView ListView2 
         Height          =   4815
         Left            =   -69880
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   11295
         _Version        =   983040
         _ExtentX        =   19923
         _ExtentY        =   8493
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListView ListView3 
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   11295
         _Version        =   983040
         _ExtentX        =   19923
         _ExtentY        =   8493
         _StockProps     =   77
         BackColor       =   -2147483643
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnRestoreD 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   5640
         Width           =   2055
         _Version        =   983040
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Restore DISC"
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnRestoreMember 
         Height          =   615
         Left            =   -69880
         TabIndex        =   5
         Top             =   5640
         Visible         =   0   'False
         Width           =   2055
         _Version        =   983040
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Restore Member"
         Enabled         =   0   'False
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnRestoreUser 
         Height          =   615
         Left            =   -69880
         TabIndex        =   6
         Top             =   5640
         Visible         =   0   'False
         Width           =   2055
         _Version        =   983040
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Restore User"
         Enabled         =   0   'False
         Appearance      =   6
      End
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRestoreD_Click()
If MsgBox("Restore DISC?", vbQuestion + vbYesNo) = vbYes Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_CDList where CD_ID like '" & Me.ListView3.SelectedItem.Text & "'", cn, 1, 2
            !Status = "AVAILABLE"
            .Update
        .Close
    End With
    Call loadDISC
    MsgBox "DISC has been retored.", vbInformation + vbOKOnly
    Me.btnRestoreD.Enabled = False
End If
End Sub

Private Sub btnRestoreMember_Click()
If MsgBox("Restore Member?", vbQuestion + vbYesNo) = vbYes Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_members where MemberID like '" & Me.ListView1.SelectedItem.Text & "'", cn, 1, 2
            !Status = "ACTIVE"
            .Update
        .Close
    End With
    Call loadMembers
    MsgBox "Member has been retored.", vbInformation + vbOKOnly
    Me.btnRestoreUser.Enabled = False
End If
End Sub

Private Sub btnRestoreUser_Click()
If MsgBox("Restore User?", vbQuestion + vbYesNo) = vbYes Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Users where Username like '" & Me.ListView2.SelectedItem.ListSubItems(2).Text & "'", cn, 1, 2
            !Status = "ACTIVE"
            .Update
        .Close
    End With
    Call loadUser
    MsgBox "User has been retored.", vbInformation + vbOKOnly
    Me.btnRestoreMember.Enabled = False
End If
End Sub

Private Sub Form_Load()
With Me.ListView1
    .ColumnHeaders.Add , , "Member ID", 2200
    .ColumnHeaders.Add , , "Member Name", 9000
End With

With Me.ListView2
    .ColumnHeaders.Add , , "Name", 5200
    .ColumnHeaders.Add , , "User Type", 4000
    .ColumnHeaders.Add , , "Username", 2000
End With

With Me.ListView3
    .ColumnHeaders.Add , , "CD ID", 1500
    .ColumnHeaders.Add , , "Title", 3000
    .ColumnHeaders.Add , , "Genre", 2000
    .ColumnHeaders.Add , , "Director", 1500
    .ColumnHeaders.Add , , "Actor", 3200
End With

Call loadMembers

Call loadUser

Call loadDISC

Me.Top = 0
Me.Left = 0
Me.TabControl1.SelectedItem = 0
End Sub


Sub loadMembers()
Me.ListView1.ListItems.Clear
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_members where status like '" & "DELETED" & "'", cn, 1, 2
        Do Until .EOF
            Set lst = Me.ListView1.ListItems.Add(, , !MemberID)
                lst.ListSubItems.Add , , !Name
            .MoveNext
        Loop
    .Close
End With
End Sub


Sub loadUser()
Me.ListView2.ListItems.Clear
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users where status like '" & "DELETED" & "'", cn, 1, 2
        Do Until .EOF
            Set lst = Me.ListView2.ListItems.Add(, , !Lname & ", " & !Fname & " " & !Mname)
                lst.ListSubItems.Add , , !UserType
                lst.ListSubItems.Add , , !UserName
                .MoveNext
        Loop
    .Close
End With
End Sub


Sub loadDISC()
Me.ListView3.ListItems.Clear
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CDLIst where status like '" & "DELETED" & "' or status like '" & "LOSE" & "' or status like '" & "DAMAGE" & "'", cn, 1, 2
        Do Until .EOF
            Set lst = Me.ListView3.ListItems.Add(, , !CD_ID)
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !Ggenre
                lst.ListSubItems.Add , , !director
                lst.ListSubItems.Add , , !actor
                
                .MoveNext
        Loop
    .Close
End With
End Sub

Private Sub ListView1_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Me.btnRestoreMember.Enabled = True
End Sub

Private Sub ListView2_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Me.btnRestoreUser.Enabled = True
End Sub

Private Sub ListView3_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Me.btnRestoreD.Enabled = True
End Sub
