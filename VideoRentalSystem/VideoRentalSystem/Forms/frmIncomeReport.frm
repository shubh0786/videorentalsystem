VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmIncomeReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   12600
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7095
      _Version        =   851972
      _ExtentX        =   12515
      _ExtentY        =   8493
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
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker3 
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
      _Version        =   851972
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   68
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4815
      Left            =   7440
      TabIndex        =   3
      Top             =   360
      Width           =   4935
      _Version        =   851972
      _ExtentX        =   8705
      _ExtentY        =   8493
      _StockProps     =   79
      Caption         =   "Total Income"
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
      Begin XtremeSuiteControls.FlatEdit txtAmount 
         Height          =   615
         Left            =   1200
         TabIndex        =   7
         Top             =   3480
         Width           =   2775
         _Version        =   851972
         _ExtentX        =   4895
         _ExtentY        =   1085
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "FlatEdit1"
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.Label lvlFrom 
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Label5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lvlTo 
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Label4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
         _Version        =   851972
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Total Amount"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   615
         _Version        =   851972
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "To"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   495
         _Version        =   851972
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "From"
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   7095
      _Version        =   851972
      _ExtentX        =   12515
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Filter By"
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
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   735
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         _Version        =   851972
         _ExtentX        =   5318
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "From"
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
         Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
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
            MinDate         =   32874
            Format          =   1
            CurrentDate     =   32874
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   735
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   3255
         _Version        =   851972
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "To"
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
         Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Width           =   2535
            _Version        =   851972
            _ExtentX        =   4471
            _ExtentY        =   661
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
            Format          =   1
         End
      End
   End
   Begin XtremeSuiteControls.PushButton bttnClose 
      Height          =   855
      Left            =   10560
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   1508
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
   End
   Begin XtremeSuiteControls.PushButton bttnPrint 
      Height          =   855
      Left            =   8760
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   1508
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
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker4 
      Height          =   735
      Left            =   7200
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
      _Version        =   851972
      _ExtentX        =   2778
      _ExtentY        =   1296
      _StockProps     =   68
      Format          =   1
   End
End
Attribute VB_Name = "frmIncomeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
Me.MousePointer = 11

    Set exlobj = CreateObject("excel.application")
    exlobj.Workbooks.Add
    With exlobj.ActiveSheet
    Dim fdte, tdte As Date
fdte = Format(Me.lvlFrom.Caption, "mm/dd/yyyy")
tdte = Format(Me.lvlTo.Caption, "mm-dd-yyyy")
        .Cells(2, 3).value = "Video Rental System  "
        .Cells(2, 3).Font.Name = "Times New Roman"
        .Cells(2, 3).Font.Size = 18
        .Cells(2, 3).Font.Bold = True
        .Cells(3, 4).value = "STI College Surigao"
        .Cells(3, 4).Font.Name = "Times New Roman"
        .Cells(3, 4).Font.Size = 14
        .Cells(3, 4).Font.Name = "Times New Roman"

        .Cells(5, 5).value = "Daily Income Report"
        .Cells(5, 5).Font.Name = "Times New Roman"
        .Cells(5, 5).Font.Size = 16
        .Cells(5, 5).Font.Name = "Times New Roman"
        .Cells(5, 5).Font.Bold = True

        .Cells(6, 4).value = "From"
        .Cells(6, 4).Font.Name = "Times New Roman"
        .Cells(6, 4).Font.Size = 12
        .Cells(6, 4).Font.Name = "Times New Roman"
        
        .Cells(6, 5).value = fdte
        .Cells(6, 5).Font.Name = "Times New Roman"
        .Cells(6, 5).Font.Size = 14
        .Cells(6, 5).Font.Name = "Times New Roman"
        .Cells(6, 5).Font.Bold = True
        
        
        .Cells(6, 7).value = "To"
        .Cells(6, 7).Font.Name = "Times New Roman"
        .Cells(6, 7).Font.Size = 12
        .Cells(6, 7).Font.Name = "Times New Roman"
        
        .Cells(6, 8).value = tdte
        .Cells(6, 8).Font.Name = "Times New Roman"
        .Cells(6, 8).Font.Size = 14
        .Cells(6, 8).Font.Name = "Times New Roman"
        .Cells(6, 8).Font.Bold = True
        
        
        '---------------------
        .Cells(9, 1).value = "Income ID"
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Size = 12
        .Cells(9, 1).Font.Name = "Times New Roman"
        .Cells(9, 1).Font.Bold = True
        
        '----------------------
        .Cells(9, 3).value = "Amount"
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Size = 12
        .Cells(9, 3).Font.Name = "Times New Roman"
        .Cells(9, 3).Font.Bold = True
        
        '-----------------------
        .Cells(9, 5).value = "Income Type"
        .Cells(9, 5).Font.Name = "Times New Roman"
        .Cells(9, 5).Font.Size = 12
        .Cells(9, 5).Font.Name = "Times New Roman"
        .Cells(9, 5).Font.Bold = True
        
        '------------------------
        .Cells(9, 7).value = "Date"
        .Cells(9, 7).Font.Name = "Times New Roman"
        .Cells(9, 7).Font.Size = 12
        .Cells(9, 7).Font.Name = "Times New Roman"
        .Cells(9, 7).Font.Bold = True
        
        '------------------------
        .Cells(9, 9).value = "User ID"
        .Cells(9, 9).Font.Name = "Times New Roman"
        .Cells(9, 9).Font.Size = 12
        .Cells(9, 9).Font.Name = "Times New Roman"
        .Cells(9, 9).Font.Bold = True
        
        
        Dim rw As Integer
        rw = 10
        Dim x As Date
        '---------------------------------------------------
        For I = 1 To Me.ListView1.ListItems.count
        Me.DateTimePicker4.value = Me.ListView1.ListItems(I).ListSubItems(3)
        x = Format(Me.DateTimePicker4.value, "mm-dd-yyyy")
            .Cells(rw, 1).value = Me.ListView1.ListItems(I)
            .Cells(rw, 1).Font.Name = "Times New Roman"
            .Cells(rw, 1).Font.Size = 12
            .Cells(rw, 1).Font.Name = "Times New Roman"
            
            
            '----------------------
            .Cells(rw, 3).value = Me.ListView1.ListItems(I).ListSubItems(1)
            .Cells(rw, 3).Font.Name = "Times New Roman"
            .Cells(rw, 3).Font.Size = 12
            .Cells(rw, 3).Font.Name = "Times New Roman"
            
            
            '----------------------
            .Cells(rw, 5).value = Me.ListView1.ListItems(I).ListSubItems(2)
            .Cells(rw, 5).Font.Name = "Times New Roman"
            .Cells(rw, 5).Font.Size = 12
            .Cells(rw, 5).Font.Name = "Times New Roman"
            
            '---------------------
            .Cells(rw, 7).value = x
            .Cells(rw, 7).Font.Name = "Times New Roman"
            .Cells(rw, 7).Font.Size = 12
            .Cells(rw, 7).Font.Name = "Times New Roman"
            
            '---------------------
            .Cells(rw, 9).value = Me.ListView1.ListItems(I).ListSubItems(4)
            .Cells(rw, 9).Font.Name = "Times New Roman"
            .Cells(rw, 9).Font.Size = 12
            .Cells(rw, 9).Font.Name = "Times New Roman"
        
            rw = rw + 1
        Next I
        
        
        '---------TOTAL AMOUNT--------------
        .Cells(39, 1).value = "Total Amount:"
        .Cells(39, 1).Font.Name = "Times New Roman"
        .Cells(39, 1).Font.Size = 16
        .Cells(39, 1).Font.Name = "Times New Roman"
        .Cells(39, 1).Font.Bold = True
        
        .Cells(39, 4).value = Me.txtAmount.Text & "   Pesos"
        .Cells(39, 4).Font.Name = "Times New Roman"
        .Cells(39, 4).Font.Size = 14
        .Cells(39, 4).Font.Name = "Times New Roman"
        .Cells(39, 4).Font.Bold = True
    End With
exlobj.Visible = True
Me.MousePointer = 0
End Sub

Private Sub DateTimePicker1_Change()
Call loadR
End Sub

Private Sub DateTimePicker2_Change()
Call loadR
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Me.DateTimePicker1.value = Now
Me.DateTimePicker2.value = Now
Me.DateTimePicker3.value = Now

With Me.ListView1
    .ColumnHeaders.Add , , "Income ID", 1000
    .ColumnHeaders.Add , , "Amount", 1000
    .ColumnHeaders.Add , , "Income Type", 2200
    .ColumnHeaders.Add , , "Date", 2000
    .ColumnHeaders.Add , , "User ID", 1000
End With

Call loadR
End Sub


Sub loadR()
Me.ListView1.ListItems.clear

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_income", cn, 1, 2
    Do Until .EOF
        Me.DateTimePicker3.value = !iDate
                    
            If Me.DateTimePicker3.value >= Me.DateTimePicker1.value And Me.DateTimePicker3.value <= Me.DateTimePicker2.value Then
                Set lst = Me.ListView1.ListItems.Add(, , !incomeID)
                    lst.ListSubItems.Add , , !amount
                    lst.ListSubItems.Add , , !iType
                    lst.ListSubItems.Add , , !iDate
                    lst.ListSubItems.Add , , !UserID
            ElseIf Me.DateTimePicker3.Year = Me.DateTimePicker1.Year And Me.DateTimePicker3.Month = Me.DateTimePicker1.Month And Me.DateTimePicker3.Day = Me.DateTimePicker1.Day And Me.DateTimePicker3.Year = Me.DateTimePicker2.Year And Me.DateTimePicker3.Month = Me.DateTimePicker2.Month And Me.DateTimePicker3.Day = Me.DateTimePicker2.Day Then
                    Set lst = Me.ListView1.ListItems.Add(, , !incomeID)
                        lst.ListSubItems.Add , , !amount
                        lst.ListSubItems.Add , , !iType
                        lst.ListSubItems.Add , , !iDate
                        lst.ListSubItems.Add , , !UserID
             ElseIf Me.DateTimePicker3.Year = Me.DateTimePicker1.Year And Me.DateTimePicker3.Month = Me.DateTimePicker1.Month And Me.DateTimePicker3.Day = Me.DateTimePicker1.Day And Me.DateTimePicker3.Year <= Me.DateTimePicker2.Year And Me.DateTimePicker3.Month <= Me.DateTimePicker2.Month And Me.DateTimePicker3.Day <= Me.DateTimePicker2.Day Then
                    Set lst = Me.ListView1.ListItems.Add(, , !incomeID)
                        lst.ListSubItems.Add , , !amount
                        lst.ListSubItems.Add , , !iType
                        lst.ListSubItems.Add , , !iDate
                        lst.ListSubItems.Add , , !UserID
            End If
        .MoveNext
        Loop
End With
Me.lvlFrom.Caption = Me.DateTimePicker1.value
Me.lvlTo.Caption = Me.DateTimePicker2.value
Call countI
End Sub

Sub countI()
Dim x  As Integer
    For I = 1 To Me.ListView1.ListItems.count
        Y = Me.ListView1.ListItems(I).SubItems(1)
        x = x + CInt(Y)
    Next I
Me.txtAmount.Text = x
End Sub

