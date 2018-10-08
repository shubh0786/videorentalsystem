VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmAuditTrail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit Trail"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9870
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      _Version        =   851972
      _ExtentX        =   16536
      _ExtentY        =   8070
      _StockProps     =   77
      BackColor       =   -2147483643
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnCreateReport 
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Create Report"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5280
      Width           =   2415
      _Version        =   851972
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
      _Version        =   851972
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker3 
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
      _Version        =   851972
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   735
      Left            =   8280
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
      _Version        =   851972
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Close"
      Appearance      =   6
   End
End
Attribute VB_Name = "frmAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public all As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnCreateReport_Click()
Me.MousePointer = 11

    Set exlobj = CreateObject("excel.application")
    exlobj.Workbooks.Add
    With exlobj.ActiveSheet
            .Cells(2, 3).value = "D-Light Computerized DVD Rental System    "
        .Cells(2, 3).Font.Name = "Times New Roman"
        .Cells(2, 3).Font.Size = 18
        .Cells(2, 3).Font.Bold = True
        .Cells(3, 4).value = "Kitcharao, Agusan del Norte"
        .Cells(3, 4).Font.Name = "Times New Roman"
        .Cells(3, 4).Font.Size = 12
        .Cells(3, 4).Font.Name = "Times New Roman"

        .Cells(5, 5).value = "Audit Trail"
        .Cells(5, 5).Font.Name = "Times New Roman"
        .Cells(5, 5).Font.Size = 16
        .Cells(5, 5).Font.Name = "Times New Roman"
        .Cells(5, 5).Font.Bold = True
        
        
         '---------------------
        .Cells(9, 1).value = "Username"
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Size = 12
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Bold = True
        
        '----------------------
        .Cells(9, 3).value = "What have done"
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Size = 12
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Bold = True
        
        
        '------------------------
        .Cells(9, 8).value = "Date & Time"
        .Cells(9, 8).Font.Name = "Times New Roman"
        .Cells(9, 8).Font.Size = 12
        .Cells(9, 8).Font.Name = "Times New Roman"
        .Cells(9, 8).Font.Bold = True
        
        Dim rw As Integer
        rw = 10
        Dim x As Date
        '---------------------------------------------------
        For I = 1 To Me.ListView1.ListItems.count
        
            .Cells(rw, 1).value = Me.ListView1.ListItems(I)
            .Cells(rw, 1).Font.Name = "Times New Roman"
            .Cells(rw, 1).Font.Size = 12
            .Cells(rw, 1).Font.Name = "Times New Roman"
            
            
            '----------------------
            .Cells(rw, 3).value = Me.ListView1.ListItems(I).ListSubItems(1)
            .Cells(rw, 3).Font.Name = "Times New Roman"
            .Cells(rw, 3).Font.Size = 12
            .Cells(rw, 3).Font.Name = "Times New Roman"
            

            '---------------------
            .Cells(rw, 8).value = Me.ListView1.ListItems(I).ListSubItems(2)
            .Cells(rw, 8).Font.Name = "Times New Roman"
            .Cells(rw, 8).Font.Size = 12
            .Cells(rw, 8).Font.Name = "Times New Roman"

            rw = rw + 1
        Next I
        End With
exlobj.Visible = True
Me.MousePointer = 0
        
    
End Sub

Private Sub DateTimePicker1_Change()
    all = False
    Call loadRecords
End Sub

Private Sub DateTimePicker2_Change()
    all = False
    Call loadRecords
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

With Me.ListView1
    .ColumnHeaders.Add , , "User ID", 1500
    .ColumnHeaders.Add , , "What have Done", 6500
    .ColumnHeaders.Add , , "Date and time", 2000
End With
Me.DateTimePicker1 = Now
Me.DateTimePicker2 = Now
all = True
Call loadRecords
End Sub


Sub loadRecords()
Me.ListView1.ListItems.clear

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_AuditTrail", cn, 1, 2
    
    If all = False Then
        Do Until .EOF
            Me.DateTimePicker3.value = !Date
                        
                If Me.DateTimePicker3.value >= Me.DateTimePicker1.value And Me.DateTimePicker3.value <= Me.DateTimePicker2.value Then
                    Set lst = Me.ListView1.ListItems.Add(, , !UserID)
                        lst.ListSubItems.Add , , !wDone
                        lst.ListSubItems.Add , , !Date
                ElseIf Me.DateTimePicker3.Year = Me.DateTimePicker1.Year And Me.DateTimePicker3.Month = Me.DateTimePicker1.Month And Me.DateTimePicker3.Day = Me.DateTimePicker1.Day And Me.DateTimePicker3.Year = Me.DateTimePicker2.Year And Me.DateTimePicker3.Month = Me.DateTimePicker2.Month And Me.DateTimePicker3.Day = Me.DateTimePicker2.Day Then
                    Set lst = Me.ListView1.ListItems.Add(, , !UserID)
                        lst.ListSubItems.Add , , !wDone
                        lst.ListSubItems.Add , , !Date
                 ElseIf Me.DateTimePicker3.Year = Me.DateTimePicker1.Year And Me.DateTimePicker3.Month = Me.DateTimePicker1.Month And Me.DateTimePicker3.Day = Me.DateTimePicker1.Day And Me.DateTimePicker3.Year <= Me.DateTimePicker2.Year And Me.DateTimePicker3.Month <= Me.DateTimePicker2.Month And Me.DateTimePicker3.Day <= Me.DateTimePicker2.Day Then
                    Set lst = Me.ListView1.ListItems.Add(, , !UserID)
                        lst.ListSubItems.Add , , !wDone
                        lst.ListSubItems.Add , , !Date
                End If
        .MoveNext
        Loop
    ElseIf all = True Then
        Do Until .EOF
            Set lst = Me.ListView1.ListItems.Add(, , !UserID)
                lst.ListSubItems.Add , , !wDone
                lst.ListSubItems.Add , , !Date
        .MoveNext
        Loop
    End If
    
End With
End Sub
