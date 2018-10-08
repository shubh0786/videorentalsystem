VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmLogHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log History"
   ClientHeight    =   5685
   ClientLeft      =   105
   ClientTop       =   885
   ClientWidth     =   12270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   12270
   Begin XtremeSuiteControls.ListView lstVW 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11655
      _Version        =   851972
      _ExtentX        =   20558
      _ExtentY        =   8281
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
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5040
      Width           =   3615
      _Version        =   851972
      _ExtentX        =   6376
      _ExtentY        =   873
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
      _Version        =   851972
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Picture         =   "frmLogHistory.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   5040
      Width           =   3495
      _Version        =   851972
      _ExtentX        =   6165
      _ExtentY        =   873
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker3 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   68
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton btnCreateRport 
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   5040
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Create Report"
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
      Picture         =   "frmLogHistory.frx":039A
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   5040
      Width           =   255
      _Version        =   851972
      _ExtentX        =   450
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "To"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   495
      _Version        =   851972
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "From"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLogHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub loadRecords()
    With lstVW
        .ColumnHeaders.clear
        .ListItems.clear
        .ColumnHeaders.Add , , "Username", 1500
        .ColumnHeaders.Add , , "Login Date", 1500
        .ColumnHeaders.Add , , "Login Time", 1500
        .ColumnHeaders.Add , , "Logout Time", 1500
        .ColumnHeaders.Add , , "Logout Date", 1500
        .ColumnHeaders.Add , , "Unit Used", 1500
    End With
    
Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_LogHistory", cn, 1, 2
        
    Do Until .EOF
    Me.DateTimePicker3.value = !LoginDate
        If Me.DateTimePicker3.Year >= Me.DateTimePicker1.Year And Me.DateTimePicker3.Year <= Me.DateTimePicker2.Year Then
            Set lst = lstVW.ListItems.Add(, , !UserName)
                lst.ListSubItems.Add , , !LoginDate
                lst.ListSubItems.Add , , !loginTime
                lst.ListSubItems.Add , , !logOutTime
                lst.ListSubItems.Add , , !logoutDate
                lst.ListSubItems.Add , , !unitUsed
        End If
        .MoveNext
    Loop
    End With
    Set rs = Nothing
End Sub

Private Sub btnCreateRport_Click()
Me.MousePointer = 11

    Set exlobj = CreateObject("excel.application")
    exlobj.Workbooks.Add
    With exlobj.ActiveSheet
    
        .Cells(2, 3).value = "Video Rental System   "
        .Cells(2, 3).Font.Name = "Times New Roman"
        .Cells(2, 3).Font.Size = 18
        .Cells(2, 3).Font.Bold = True
        .Cells(3, 4).value = "STI Collge Surigao"
        .Cells(3, 4).Font.Name = "Times New Roman"
        .Cells(3, 4).Font.Size = 12
        .Cells(3, 4).Font.Name = "Times New Roman"

        .Cells(5, 4).value = "LOG HISTORY"
        .Cells(5, 4).Font.Name = "Times New Roman"
        .Cells(5, 4).Font.Size = 16
        .Cells(5, 4).Font.Name = "Times New Roman"
        .Cells(5, 4).Font.Bold = True
        
         '---------------------
        .Cells(9, 1).value = "Username"
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Size = 12
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Bold = True
        
        '----------------------
        .Cells(9, 3).value = "Login Date & Time"
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Size = 12
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Bold = True
        
        
        '------------------------
        .Cells(9, 6).value = "Logout Date & Time"
        .Cells(9, 6).Font.Name = "Times New Roman"
        .Cells(9, 6).Font.Size = 12
        .Cells(9, 6).Font.Name = "Times New Roman"
        .Cells(9, 6).Font.Bold = True
        
        '------------------------
        .Cells(9, 8).value = "Unit used"
        .Cells(9, 8).Font.Name = "Times New Roman"
        .Cells(9, 8).Font.Size = 12
        .Cells(9, 8).Font.Name = "Times New Roman"
        .Cells(9, 8).Font.Bold = True
        
        
        
        Dim rw As Integer
        rw = 10
        Dim X As Date
        '---------------------------------------------------
        For I = 1 To Me.lstVW.ListItems.count
        
            .Cells(rw, 1).value = Me.lstVW.ListItems(I)
            .Cells(rw, 1).Font.Name = "Times New Roman"
            .Cells(rw, 1).Font.Size = 12
            .Cells(rw, 1).Font.Name = "Times New Roman"
            
            
     
            .Cells(rw, 3).value = Me.lstVW.ListItems(I).ListSubItems(1) & " " & Me.lstVW.ListItems(I).ListSubItems(2)
            .Cells(rw, 3).Font.Name = "Times New Roman"
            .Cells(rw, 3).Font.Size = 12
            .Cells(rw, 3).Font.Name = "Times New Roman"
            
            
 
            .Cells(rw, 6).value = Me.lstVW.ListItems(I).ListSubItems(3) & " " & Me.lstVW.ListItems(I).ListSubItems(4)
            .Cells(rw, 6).Font.Name = "Times New Roman"
            .Cells(rw, 6).Font.Size = 12
            .Cells(rw, 6).Font.Name = "Times New Roman"
            

            .Cells(rw, 8).value = Me.lstVW.ListItems(I).ListSubItems(5)
            .Cells(rw, 8).Font.Name = "Times New Roman"
            .Cells(rw, 8).Font.Size = 12
            .Cells(rw, 8).Font.Name = "Times New Roman"

            rw = rw + 1
        Next I
        End With
exlobj.Visible = True
Me.MousePointer = 0
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub DateTimePicker1_Change()
Call loadRecords
End Sub

Private Sub DateTimePicker2_Change()
Call loadRecords
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Me.DateTimePicker1.value = Now
Me.DateTimePicker2.value = Now
Call loadRecords
End Sub
