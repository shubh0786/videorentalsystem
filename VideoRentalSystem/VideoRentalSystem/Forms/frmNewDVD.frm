VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmNewDVD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NEW CD ENTRY"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lvView 
      Height          =   3615
      Left            =   5280
      TabIndex        =   22
      Top             =   240
      Width           =   6615
      _Version        =   851972
      _ExtentX        =   11668
      _ExtentY        =   6376
      _StockProps     =   77
      BackColor       =   -2147483643
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   5280
      TabIndex        =   24
      Top             =   3960
      Width           =   6615
      _Version        =   851972
      _ExtentX        =   11668
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Search"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtSby 
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   360
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboSearch 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSearch 
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   360
         Width           =   2655
         _Version        =   851972
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Search By:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.PushButton bttnClear 
      Height          =   975
      Left            =   840
      TabIndex        =   23
      Top             =   4080
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton bttnADD 
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.PushButton bttnCancel 
      Height          =   975
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
      _Version        =   851972
      _ExtentX        =   1931
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      TextImageRelation=   1
   End
   Begin XtremeSuiteControls.PushButton bttnGenre 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   495
      _Version        =   851972
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtDirector 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtTitle 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtDVDNo 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAct 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtDesc 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtGenre 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
      _Version        =   851972
      _ExtentX        =   4260
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cboGenre 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
      _Version        =   851972
      _ExtentX        =   4683
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtYearR 
      Height          =   315
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   3015
      _Version        =   851972
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboRelease 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCDType 
      Height          =   315
      Left            =   1680
      TabIndex        =   21
      Top             =   3480
      Width           =   3015
      _Version        =   851972
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboCDType 
      Height          =   315
      Left            =   1680
      TabIndex        =   18
      Top             =   3480
      Width           =   3255
      _Version        =   851972
      _ExtentX        =   5741
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CD Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   240
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CD Number:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Description:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Year Realease:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actor/Actress"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Director:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Genre:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Title:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmNewDVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bttnADD_Click()
'*************************
'   RECORD TRANSACTION to tbl_Transaction
'*************************
Call transaction("Add New CD")


'**************************
'   RECORD DVD on tbl_CDList
'**************************
Set rs = New ADODB.Recordset
    With rs
        .Open "select * from tbl_CDList", cn, 1, 2
        
        .AddNew
        !CDNo = Me.txtDVDNo.Text
        !Title = Me.txtTitle.Text
        !GGenre = Me.txtGenre.Text
        !director = Me.txtDirector.Text
        !actor = Me.txtAct.Text
        !YearRelease = Me.cboRelease
        !Description = Me.txtDesc.Text
        !Status = "Available"
        !TransID = ctr
        !CDType = Me.cboCDType.Text
        !goodFor = "One(1) Day"
        .Update

        Call addGenre
        Call addDirector
        Call addActor
        If MsgBox("New Entry has been Successfully Added.", vbInformation + vbOKOnly) = vbOK Then clearF
            Call cboItems
            If dvdList = True Then
                frmDVD.loadRecords
            End If
            Call clearF
    End With
    Set rs = Nothing
    Call loadRecords
End Sub

Sub getCount()
'************************************
'       Get CD Count
'************************************
Dim rctr  As Integer
rctr = 1
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CDList", cn, 1, 2
    
    Do Until .EOF
    rctr = rctr + 1
    
    .MoveNext
    Loop
    txtDVDNo.Text = rctr
End With
Set rs = Nothing
End Sub


Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnClear_Click()
    Call clearF
End Sub

Private Sub bttnGenre_Click()
Call cboItems
txtGenre.Text = ""
End Sub

Private Sub cboCDType_Click()
    Me.txtCDType.Locked = False
    Me.txtCDType.Text = Me.cboCDType.Text
    Me.txtCDType.Locked = True
End Sub

Private Sub cboGenre_Change()
txtGenre.Text = txtGenre.Text & "/" & cboGenre.Text
End Sub

Private Sub cboGenre_Click()
On Error GoTo err_Handler
If txtGenre.Text = "" Then
    txtGenre.Text = txtGenre.Text & cboGenre.Text
    cboGenre.RemoveItem (cboGenre.ListIndex)
Else
    txtGenre.Text = txtGenre.Text & "/" & cboGenre.Text
    cboGenre.RemoveItem (cboGenre.ListIndex)
End If
Exit Sub
err_Handler:
    MsgBox err.Description, vbCritical + vbOKOnly
    Call cboItems
    Me.txtGenre.Text = ""
End Sub

Private Sub cboRelease_Click()
Me.txtYearR.Locked = False
Me.txtYearR.Text = Me.cboRelease.Text
Me.txtYearR.Locked = True
End Sub

Private Sub cboSearch_Click()
Me.txtSby.Locked = False
Me.txtSby.Text = Me.cboSearch.Text
Me.txtSby.Locked = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub


Private Sub Form_Load()
    Call clearF
    Call cboItems
    
    Call loadRecords
    Call SearchBy
End Sub

Sub SearchBy()
    Me.cboSearch.AddItem "CD Number"
    Me.cboSearch.AddItem "Title"
    
    Me.txtSby.Text = "CD Number"
End Sub
Sub checkFields()
With Me
    If .txtTitle.Text = "" Or .txtGenre.Text = "" Or .txtAct.Text = "" Or .txtDirector.Text = "" Or .cboCDType.Text = "" Or .cboRelease.Text = "" Then
        .bttnADD.Enabled = False
    Else
        .bttnADD.Enabled = True
    End If
End With
End Sub


Sub cboItems()
'****************************************************
'       GENRE
'****************************************************
    cboGenre.clear

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Genre", cn, 1, 2
    
    If .EOF = False Then
        Do Until .EOF
            cboGenre.AddItem (!genre)
            .MoveNext
        Loop
    End If
End With
'***************************************************
'       YEAR RELEASE
'***************************************************
Dim ctr As Integer
ctr = 2011
Do Until ctr = 1990
    cboRelease.AddItem (ctr)
    If ctr = 1990 Then Exit Do
    ctr = ctr - 1
Loop
    
'**************************************************
'       CD Type
'**************************************************
With Me
    .cboCDType.clear
    .cboCDType.AddItem "DVD"
    .cboCDType.AddItem "VCD"
End With
End Sub

Sub clearF()
    'txtDVDNo.Text = ""
    txtTitle.Text = ""
    cboGenre.Text = ""
    txtDirector.Text = ""
    txtAct.Text = ""
    cboRelease.Text = ""
    txtDesc.Text = ""
    Me.txtGenre.Text = ""
    Me.txtCDType.Text = ""
    Me.txtYearR.Text = ""
    Call getCount
    Me.bttnADD.Enabled = False
    bttnGenre_Click
End Sub

Private Sub txtAct_Change()
    Call checkFields
End Sub

Private Sub txtCDType_Change()
    Call checkFields
End Sub

Private Sub txtDirector_Change()
    Call checkFields
End Sub

Private Sub txtGenre_Change()
    Call checkFields
End Sub

Private Sub txtGenre_KeyPress(KeyAscii As Integer)
Dim letter As String
letter = KeyAscii
Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_Genre where Genre like '" & KeyAscii & "%'", cn, 1, 2
    End With
End Sub


Sub addGenre()
Dim genre1() As String
Dim sPlitCtr As Integer

Set rs = New ADODB.Recordset
    With rs
        genre1 = Split(Me.txtGenre.Text, "/")
        For sPlitCtr = 0 To UBound(genre1)

        .Open "Select * from tbl_Genre", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !genre = genre1(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If UCase(genre1(sPlitCtr)) = UCase(!genre) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !genre = genre1(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub

Sub addDirector()
Dim director() As String
Dim sPlitCtr As Integer

Set rs = New ADODB.Recordset
    With rs
        director = Split(Me.txtDirector.Text, "/")
        For sPlitCtr = 0 To UBound(director)

        .Open "Select * from tbl_Director", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !DirectorName = director(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If Trim(UCase(director(sPlitCtr))) = Trim(UCase(!DirectorName)) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !DirectorName = director(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub

Sub addActor()
Dim actor() As String
Dim sPlitCtr As Integer

Set rs = New ADODB.Recordset
    With rs
        actor = Split(Me.txtAct.Text, "/")
        For sPlitCtr = 0 To UBound(actor)

        .Open "Select * from tbl_Actor_Actress", cn, 1, 2
            If .EOF = True Then
                    .AddNew
                    !ActorActressName = actor(sPlitCtr)
                    .Update
                    rs.Close
            Else
                Do Until .EOF
                    If Trim(UCase(actor(sPlitCtr))) = Trim(UCase(!ActorActressName)) Then
                        .Close
                        Exit Do
                    Else
                        .MoveNext
                        If .EOF = True Then
                            .AddNew
                            !ActorActressName = actor(sPlitCtr)
                            .Update
                            rs.Close
                            Exit Do
                        End If
                    End If
                Loop
            End If
        Next sPlitCtr
    End With
End Sub


Private Sub txtSearch_Change()
Set rs = New ADODB.Recordset
With rs
    If Me.txtSby.Text = "" Or Me.txtSby.Text = "CD Number" Then
        .Open "Select * from tbl_CdList where CDNo like '" & Me.txtSearch.Text & "%'", cn, 1, 2
    ElseIf Me.txtSby.Text = "Title" Then
        .Open "Select * from tbl_CDList where Title like '" & Me.txtSearch.Text & "%'", cn, 1, 2
    End If
    Me.lvView.ListItems.clear
        Do Until .EOF
            Set lst = lvView.ListItems.Add(, , !CDNo)
            lst.ListSubItems.Add , , !Title
            lst.ListSubItems.Add , , !GGenre
            lst.ListSubItems.Add , , !director
            lst.ListSubItems.Add , , !actor
            lst.ListSubItems.Add , , !YearRelease
            lst.ListSubItems.Add , , !Description
            lst.ListSubItems.Add , , !Status
            lst.ListSubItems.Add , , !CDType
            .MoveNext
        Loop
    
    .Close
End With
End Sub

Private Sub txtTitle_Change()
    Call checkFields
End Sub

Private Sub txtYearR_Change()
    Call checkFields
End Sub


Sub loadRecords()
With lvView
    .ColumnHeaders.clear
    .ListItems.clear
    .ColumnHeaders.Add , , "CD Number", 1500
    .ColumnHeaders.Add , , "Title", 1500
    .ColumnHeaders.Add , , "Genre", 2000
    .ColumnHeaders.Add , , "Director", 1500
    .ColumnHeaders.Add , , "Actor/Actress", 1500
    .ColumnHeaders.Add , , "YearRelease", 2500
    .ColumnHeaders.Add , , "Description", 1500
    .ColumnHeaders.Add , , "Status", 1500
    .ColumnHeaders.Add , , "CD TYPE", 1500
End With

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_CdList", cn, 1, 2
    
    Do Until .EOF
        Set lst = lvView.ListItems.Add(, , !CD_ID)
        lst.ListSubItems.Add , , !Title
        lst.ListSubItems.Add , , !GGenre
        lst.ListSubItems.Add , , !director
        lst.ListSubItems.Add , , !actor
        lst.ListSubItems.Add , , !YearRelease
        lst.ListSubItems.Add , , !Category
        lst.ListSubItems.Add , , !goodFor
        lst.ListSubItems.Add , , !CDType
        .MoveNext
    Loop
    .Close
End With
End Sub
