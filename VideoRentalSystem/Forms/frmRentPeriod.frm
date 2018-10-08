VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmRentPeriod 
   Caption         =   "Rent Period"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   13140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   3375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      _Version        =   851972
      _ExtentX        =   18230
      _ExtentY        =   5953
      _StockProps     =   77
      BackColor       =   -2147483643
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnCancel 
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Picture         =   "frmRentPeriod.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
      _Version        =   851972
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Picture         =   "frmRentPeriod.frx":039A
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   720
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "To:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   855
      _Version        =   851972
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "From"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   3015
      _Version        =   851972
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
      _Version        =   851972
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Rent Period"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboPeriod 
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
      _Version        =   851972
      _ExtentX        =   4048
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker2 
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   720
      Width           =   3015
      _Version        =   851972
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _Version        =   851972
      _ExtentX        =   23310
      _ExtentY        =   10610
      _StockProps     =   79
      Appearance      =   6
   End
End
Attribute VB_Name = "frmRentPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Date
Dim Y As Date
Dim z As Date
Dim sDate, xDate As String
Sub loadStartDate()

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users", cn, 1, 2
    
        Do Until .EOF
        xDate = !dte
        Me.xDate.value = xDate
            If sDate = "" Then
                sDate = xDate
                Me.sDate.value = sDate
            End If
            Y = Format(Me.sDate.value, "mm/dd/yyyy")
            z = Format(Me.xDate.value, "mm/dd/yyyy")
        If Y > z Then
            Me.sDate.value = Me.xDate.value
        End If
            .MoveNext
    Loop
End With
End Sub

Sub loadRecords()
Me.ListView1.ListItems.clear
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CDLIst", cn, 1, 2
    
    Do Until .EOF
    Me.dAcq.value = !dateAcquired
        If Me.dAcq.value >= Me.DateTimePicker1.value And (Me.dAcq.value <= Me.DateTimePicker2.value) Then
            Set lst = Me.ListView1.ListItems.Add(, , !CD_ID)
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !Ggenre
                lst.ListSubItems.Add , , !director
                lst.ListSubItems.Add , , !actor
                lst.ListSubItems.Add , , !YearRelease
                lst.ListSubItems.Add , , !CDType
                lst.ListSubItems.Add , , !goodFor
                lst.ListSubItems.Add , , !Category
                lst.ListSubItems.Add , , !dateAcquired
        End If
        .MoveNext
    Loop
    
End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
Dim vPeriod As Integer
If Me.cboPeriod.Text = "1 DAY" Then
    vPeriod = 1
ElseIf Me.cboPeriod.Text = "2 DAYS" Then
    vPeriod = 2
ElseIf Me.cboPeriod.Text = "3 DAYS" Then
    vPeriod = 3
End If

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CDList", cn, 1, 2
    
    
    For I = 1 To Me.ListView1.ListItems.count
        Do Until .EOF
            If !CD_ID = Me.ListView1.ListItems(I) Then
                    !goodFor = vPeriod
                    .Update
                    Exit Do
            Else
                .MoveNext
            End If
        Loop
    Next I
    .Close
    MsgBox "RENT Period Change.", vbInformation + vbOKOnly
        Call loadRecords
        Me.cboPeriod.Text = ""
End With
End Sub

Private Sub DateTimePicker1_Change()
    X = Format(Me.DateTimePicker1.value, "mm/dd/yyyy")
    Call loadRecords
End Sub

Private Sub DateTimePicker2_Change()
    Call loadRecords
End Sub

Private Sub Form_Load()
   
    
    With Me.ListView1
        .ColumnHeaders.Add , , "CD ID", 1500
        .ColumnHeaders.Add , , "Title", 1500
        .ColumnHeaders.Add , , "Genre", 1000
        .ColumnHeaders.Add , , "Director", 1300
        .ColumnHeaders.Add , , "Actor/Actress", 1500
        .ColumnHeaders.Add , , "Year Release", 1000
        .ColumnHeaders.Add , , "CD Type", 1000
        .ColumnHeaders.Add , , "Rent Period(days)", 1500
        .ColumnHeaders.Add , , "Category", 1000
        .ColumnHeaders.Add , , "Date Acquired", 1500
    End With
         Call loadStartDate
        Me.DateTimePicker1.MinDate = Me.sDate.value
        Me.DateTimePicker2.MinDate = Me.sDate.value
        Me.DateTimePicker2.MaxDate = Now
        Me.DateTimePicker1.value = Me.sDate.value
        Me.DateTimePicker2.value = Now
    
    
    
    With Me.cboPeriod
        .AddItem "1 DAY"
        .AddItem "2 DAYS"
        .AddItem "3 DAYS"
        .Text = ""
    End With
    
    X = Format(Me.DateTimePicker1.value, "mm/dd/yyyy")
    
    Call loadRecords
End Sub

