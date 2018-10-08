VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CO6204~1.OCX"
Begin VB.Form frmReturn 
   Caption         =   "Return Disc"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   12735
      _Version        =   851972
      _ExtentX        =   22463
      _ExtentY        =   15055
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      Begin XtremeSuiteControls.ListView ListView1 
         Height          =   6015
         Left            =   7440
         TabIndex        =   11
         Top             =   360
         Width           =   5055
         _Version        =   851972
         _ExtentX        =   8916
         _ExtentY        =   10610
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
      Begin XtremeSuiteControls.ListView ListView2 
         Height          =   2775
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   5775
         _Version        =   851972
         _ExtentX        =   10186
         _ExtentY        =   4895
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   735
         Left            =   7440
         TabIndex        =   12
         Top             =   6600
         Width           =   5175
         _Version        =   851972
         _ExtentX        =   9128
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Filter"
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
         Begin XtremeSuiteControls.FlatEdit txtSearch 
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            Top             =   240
            Width           =   2415
            _Version        =   851972
            _ExtentX        =   4260
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
         End
         Begin VB.ComboBox cboSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   13
            Text            =   "Combo1"
            Top             =   240
            Width           =   2295
         End
      End
      Begin XtremeSuiteControls.PushButton bttnAdditional 
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   7440
         Width           =   2535
         _Version        =   851972
         _ExtentX        =   4471
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Additional Penalty"
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
         Picture         =   "frmReturn.frx":0000
      End
      Begin XtremeSuiteControls.PushButton bttnRemove 
         Height          =   615
         Left            =   3240
         TabIndex        =   8
         Top             =   7440
         Width           =   2655
         _Version        =   851972
         _ExtentX        =   4683
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Remove Item"
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
         Picture         =   "frmReturn.frx":039A
      End
      Begin XtremeSuiteControls.FlatEdit txtPenalty 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   4440
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
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtAdditional 
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   4920
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
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
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   5400
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPayment 
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   6120
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtChange 
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   6600
         Width           =   3495
         _Version        =   851972
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton bttnReturn 
         Height          =   615
         Left            =   6240
         TabIndex        =   9
         Top             =   7440
         Width           =   2775
         _Version        =   851972
         _ExtentX        =   4895
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Return"
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
         Picture         =   "frmReturn.frx":04FC
      End
      Begin XtremeSuiteControls.PushButton bttnClose 
         Height          =   615
         Left            =   9360
         TabIndex        =   10
         Top             =   7440
         Width           =   2775
         _Version        =   851972
         _ExtentX        =   4895
         _ExtentY        =   1085
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
         Picture         =   "frmReturn.frx":0896
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   6600
         Width           =   855
         _Version        =   851972
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Change:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   1920
         TabIndex        =   18
         Top             =   6120
         Width           =   975
         _Version        =   851972
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Payment:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   5400
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Total Payment:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   1560
         TabIndex        =   16
         Top             =   4920
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Other Penalty:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   4440
         Width           =   1335
         _Version        =   851972
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Due Penalty:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bP As Integer

Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub bttnRemove_Click()
Set lst = Me.ListView1.ListItems.Add(, , Me.ListView2.SelectedItem)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(1)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(2)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(3)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(4)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(5)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(6)
    lst.ListSubItems.Add , , Me.ListView2.SelectedItem.ListSubItems(7)
    
    If DateDiff("d", Me.ListView2.SelectedItem.ListSubItems(5).Text, Now) > 0 Then
        For I = 1 To 7
            lst.ListSubItems.Item(I).ForeColor = &HFF&
        Next I
    End If
    Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) - Me.ListView2.SelectedItem.ListSubItems(9).Text
    Me.txtPenalty.Text = CInt(Me.txtPenalty.Text) - Me.ListView2.SelectedItem.ListSubItems(8).Text
    Me.txtTotal.Text = CInt(Me.txtPenalty.Text) + CInt(Me.txtAdditional.Text)
    X = Me.ListView2.SelectedItem.Index
    Me.ListView2.ListItems.Remove (X)
    
    Me.bttnAdditional.Enabled = False
    Me.bttnRemove.Enabled = False
End Sub

Private Sub bttnReturn_Click()
Call getIncomeID
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_loan where status like '" & "Rent" & "'", cn, 1, 2
    
    For I = 1 To Me.ListView2.ListItems.count
    If .EOF = False Then .MoveFirst
        Set lst = Me.ListView2.ListItems.Item(I)
        Do Until .EOF
        
            If !Ino = lst.ListSubItems.Item(2) Then
                  

                If lst.ListSubItems.Item(8).Text <> "0" Then !Duefee = True
                If lst.ListSubItems.Item(10).Text = "DVD Damage" Or lst.ListSubItems.Item(10).Text = "VCD Damage" Then !damageFee = True
                If lst.ListSubItems.Item(10).Text = "DVD Lose" Or lst.ListSubItems.Item(10).Text = "VCD Lose" Then !LoseFee = True
                If lst.ListSubItems.Item(10).Text = "DVD Tampering" Or lst.ListSubItems.Item(10).Text = "VCD Tampering" Then !temperedFee = True
                
                !Status = "Return"
                !rincomeID = incomeID
                .Update

                Set rs = New ADODB.Recordset
                With rs
                    xx = "AVAILABLE"
                        If lst.ListSubItems.Item(10).Text = "DVD Damage" Or lst.ListSubItems.Item(10).Text = "VCD Damage" Then xx = "DAMAGE"
                        If lst.ListSubItems.Item(10).Text = "DVD Lose" Or lst.ListSubItems.Item(10).Text = "VCD Lose" Then xx = "LOSE"
                        If lst.ListSubItems.Item(10).Text = "DVD Tampering" Or lst.ListSubItems.Item(10).Text = "VCD Tampering" Then xx = "TAMPERED"
                        .Open "Select * from tbl_IndividualCDList where Ino like '" & lst.ListSubItems.Item(2) & "'", cn, 1, 2
                        
                        !Status = xx
                        .Update
                        .Close
                End With
                
                '***********Audit Trail
                wDone = "Return of CD" & "/" & "Loan ID" & " - " & Me.ListView2.ListItems.Item(I) & "/" & "CD Number" & " - " & lst.ListSubItems(2)
                Call auditT
            
                Exit Do
            Else
                .MoveNext
            End If
            
        Loop

    Next I
    
    .Close
    '****************************************
    '   RECORD INCOME
    '****************************************
    
    Set rs = New ADODB.Recordset
        .Open "Select * from tbl_Income", cn, 1, 2
            .AddNew
            !incomeID = incomeID
            !amount = Me.txtTotal.Text
            !iDate = Now
            !iType = "RETURN OF CD"
            !UserID = user
        .Update
        .Close

    MsgBox "Returned!", vbInformation + vboknly

    Call loadRecords
    With Me
        .ListView2.ListItems.clear
        .txtAdditional.Text = ""
        .txtChange.Text = ""
        .txtPayment.Text = ""
        .txtTotal.Text = ""
        .txtPenalty.Text = ""
        .bttnReturn.Enabled = False
    End With
End With


End Sub

Private Sub Form_Load()

    Me.Top = 1700
    Me.Left = 3200
With Me.ListView1
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "", 1
    .ColumnHeaders.Add , , "Loan ID", 1500
    .ColumnHeaders.Add , , "CD Number", 1500
    .ColumnHeaders.Add , , "Title", 2000
    .ColumnHeaders.Add , , "Rent Date", 2000
    .ColumnHeaders.Add , , "Return Date", 1500
    .ColumnHeaders.Add , , "Member ID", 1500
    .ColumnHeaders.Add , , "Member Name", 2000
End With

With Me.ListView2
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "", 1
    .ColumnHeaders.Add , , "Loan ID", 1500
    .ColumnHeaders.Add , , "CD Number", 1500
    .ColumnHeaders.Add , , "Title", 2000
    .ColumnHeaders.Add , , "Rent Date", 2000
    .ColumnHeaders.Add , , "Return Date", 1500
    .ColumnHeaders.Add , , "Member ID", 1500
    .ColumnHeaders.Add , , "Member Name", 2000
    .ColumnHeaders.Add , , "Penalty", 1500
    .ColumnHeaders.Add , , "Additional Penalty", 2000
    .ColumnHeaders.Add , , "Penalty Type", 1500
End With
With Me
    .cboSearch.AddItem "Title"
    .cboSearch.AddItem "CD Number"
    .cboSearch.AddItem "Member ID"
    
    .cboSearch.Text = "Title"
    .txtSearch.Text = ""
    .txtAdditional.Text = "0"
    .txtPenalty.Text = "0"
    .txtTotal.Text = "0"
    .bttnAdditional.Enabled = False
End With
Call loadRecords
End Sub


Sub loadRecords()
Me.ListView1.ListItems.clear
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from QUnreturn where status like 'Rent'", cn, 1, 2
    
    If .EOF = True Then
    
    Else
        Do Until .EOF
            Set lst = Me.ListView1.ListItems.Add(, , "")
                lst.ListSubItems.Add , , !loanID
                lst.ListSubItems.Add , , !Ino
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !rentDate
                lst.ListSubItems.Add , , !ReturnDate
                lst.ListSubItems.Add , , !MemberID
                lst.ListSubItems.Add , , !Lname & ", " & !Fname & " " & !Mname
                    
                    If DateDiff("d", !ReturnDate, Now) > 0 Then
                        For I = 1 To 7
                            lst.ListSubItems.Item(I).ForeColor = &HFF&
                        Next I
                    End If
            .MoveNext
        Loop
    End If
    .Close
End With
End Sub

Private Sub ListView1_DblClick()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from QUnreturn where iNO like '" & Me.ListView1.SelectedItem.ListSubItems(2) & "'", cn, 1, 2
            bCDType = !CDType
            
        If DateDiff("d", !ReturnDate, Now) > 0 Then
            bPenalty = True
            Call calcDue(DateDiff("d", !ReturnDate, Now))
        Set lst = Me.ListView2.ListItems.Add(, , "")
            lst.ListSubItems.Add , , !loanID
            lst.ListSubItems.Add , , !Ino
            lst.ListSubItems.Add , , !Title
            lst.ListSubItems.Add , , !rentDate
            lst.ListSubItems.Add , , !ReturnDate
            lst.ListSubItems.Add , , !MemberID
            lst.ListSubItems.Add , , !Lname & ", " & !Fname & " " & !Mname
            lst.ListSubItems.Add , , bP
            lst.ListSubItems.Add , , "0"
            lst.ListSubItems.Add , , "none"
        Else
            bPenalty = False
        Set lst = Me.ListView2.ListItems.Add(, , "")
            lst.ListSubItems.Add , , !loanID
            lst.ListSubItems.Add , , !Ino
            lst.ListSubItems.Add , , !Title
            lst.ListSubItems.Add , , !rentDate
            lst.ListSubItems.Add , , !ReturnDate
            lst.ListSubItems.Add , , !MemberID
            lst.ListSubItems.Add , , !Lname & ", " & !Fname & " " & !Mname
            lst.ListSubItems.Add , , "0"
            lst.ListSubItems.Add , , "0"
            lst.ListSubItems.Add , , "none"
            
            
        End If
        X = Me.ListView1.SelectedItem.Index
        Me.ListView1.ListItems.Remove (X)
    .Close
End With

Call check
End Sub

Sub calcDue(ByVal a As Integer)
Set rs = New ADODB.Recordset
With rs
    If bCDType = "DVD" Then
        .Open "Select DVDDUe from tbl_PenaltyRate", cn, 1, 2
        Me.txtPenalty.Text = CInt(Me.txtPenalty.Text) + (a * !dvdDue)
        bP = a * !dvdDue
        Me.txtTotal.Text = CInt(Me.txtTotal.Text) + bP
    ElseIf bCDType = "VCD" Then
        .Open "Select VCDDUE from tbl_Penaltyrate", cn, 1, 2
        Me.txtPenalty.Text = CInt(Me.txtPenalty.Text) + (a * !vcdDue)
        bP = a * !vcdDue
        Me.txtTotal.Text = CInt(Me.txtTotal.Text) + bP
    End If
    .Close
End With
End Sub

Sub calcAdditional()
Dim amount As Integer
Dim dType As String
amount = 0
dType = ""
On Error GoTo errH
Set rs = New ADODB.Recordset
With rs
        .Open "Select * from tbl_CDList where title like '" & Me.ListView2.SelectedItem.ListSubItems(3) & "'", cn, 1, 2
            bCDType = !CDType
        .Close
    
X = CInt(Me.txtPenalty.Text)
    If oDamage = True Then
    Set rs = New ADODB.Recordset
        If bCDType = "DVD" Then
            .Open "Select DVDDamage from tbl_PenaltyRate", cn, 1, 2
                amount = !dvdDamage
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !dvdDamage
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "DVD Damage"
            .Close
        ElseIf bCDType = "VCD" Then
            .Open "Select VCDDamage from tbl_PenaltyRate", cn, 1, 2
                amount = !vcdDamage
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !vcdDamage
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "VCD Damage"
            .Close
        End If
        X = Me.ListView2.SelectedItem.Index
        lst = Me.ListView2.ListItems.Item(X)
        lst.ListSubItems(9) = amount
        lst.ListSubItems(10) = dType
    ElseIf oLose = True Then
    Set rs = New ADODB.Recordset
        If bCDType = "DVD" Then
            .Open "Select DVDLose from tbl_PenaltyRate", cn, 1, 2
                amount = !dvdLose
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !dvdLose
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "DVD Lose"
            .Close
        ElseIf bCDType = "VCD" Then
            .Open "Select VCDLose from tbl_PenaltyRate", cn, 1, 2
                amount = !vcdLose
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !vcdLose
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "VCD Lose"
            .Close
        End If
        Me.ListView2.SelectedItem.ListSubItems(9) = amount
        Me.ListView2.SelectedItem.ListSubItems(10) = dType
    ElseIf oTampering = True Then
    Set rs = New ADODB.Recordset
        If bCDType = "DVD" Then
            .Open "Select DVDtampering from tbl_PenaltyRate", cn, 1, 2
                amount = !dvdTampering
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !dvdTampering
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "DVD Tampering"
            .Close
        ElseIf bCDType = "VCD" Then
            .Open "Select VCDtampering from tbl_PenaltyRate", cn, 1, 2
               amount = !vcdTampering
                Me.txtAdditional.Text = CInt(Me.txtAdditional.Text) + !vcdTampering
                Me.txtTotal.Text = X + CInt(Me.txtAdditional.Text)
                dType = "VCD Tampering"
            .Close
        End If
        Me.ListView2.SelectedItem.ListSubItems(9) = amount
        Me.ListView2.SelectedItem.ListSubItems(10) = dType
    End If
    
    
End With
Exit Sub
errH:
    MsgBox err.Description
End Sub


Private Sub ListView2_DblClick()
On Error GoTo errH
    X = Me.ListView2.SelectedItem
    Me.bttnAdditional.Enabled = True
    Me.bttnRemove.Enabled = True
Exit Sub
errH:
    MsgBox err.Description '"No Selected Item", vbInformation + vbOKOnly
End Sub

Private Sub ListView2_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Me.bttnRemove.Enabled = True
Me.bttnAdditional.Enabled = True
End Sub



Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.bttnAdditional.Enabled = True
End Sub

Private Sub txtAdditional_Change()
If Me.txtAdditional.Text = "" Or Me.txtPenalty.Text = "" Or Me.txtTotal.Text = "" Then
    Me.txtAdditional.Text = "0"
    Me.txtPenalty.Text = "0"
    Me.txtTotal.Text = "0"
End If

End Sub

Private Sub txtPayment_Change()
If Me.txtPayment.Text = "" Then Me.txtPayment.Text = "0"
If Me.txtChange.Text = "" Then Me.txtChange.Text = "0"
If CInt(Me.txtPayment.Text) >= CInt(Me.txtTotal.Text) Then
    Me.bttnReturn.Enabled = True
     Me.txtChange.Text = CInt(Me.txtPayment.Text) - CInt(Me.txtTotal.Text)
Else
    Me.bttnReturn.Enabled = False
End If
If Me.txtPayment.Text = "0" Then Me.txtChange.Text = "0"
If Me.txtPayment.Text = "0" Or Me.txtPayment.Text = "" Or CInt(Me.txtPayment.Text) < CInt(Me.txtTotal.Text) Then
    Me.bttnReturn.Enabled = False
    Me.txtChange.Text = "0"
End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) = True Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtPenalty_Change()
If Me.txtAdditional.Text = "" Or Me.txtPenalty.Text = "" Or Me.txtTotal.Text = "" Then
    Me.txtAdditional.Text = "0"
    Me.txtPenalty.Text = "0"
    Me.txtTotal.Text = "0"
End If

End Sub

Private Sub txtSearch_Change()
If Me.txtSearch.Text = "" Then
    Call loadRecords
Else
Me.ListView1.ListItems.clear
Set rs = New ADODB.Recordset
With rs
    If Me.cboSearch.Text = "Title" Then .Open "Select * from QUnreturn where title like '" & Me.txtSearch.Text & "%'", cn, 1, 2
    If Me.cboSearch.Text = "CD Number" Then .Open "Select * from QUnreturn where Ino like '" & Me.txtSearch.Text & "%'", cn, 1, 2
    If Me.cboSearch.Text = "Member ID" Then .Open "Select * from QUnreturn where memberID like '" & Me.txtSearch.Text & "%'", cn, 1, 2
    
    If .EOF = True Then
    
    Else
        Do Until .EOF
            Set lst = Me.ListView1.ListItems.Add(, , "")
                lst.ListSubItems.Add , , !loanID
                lst.ListSubItems.Add , , !Ino
                lst.ListSubItems.Add , , !Title
                lst.ListSubItems.Add , , !rentDate
                lst.ListSubItems.Add , , !ReturnDate
                lst.ListSubItems.Add , , !MemberID
                lst.ListSubItems.Add , , !Name
                    
                    If DateDiff("d", !ReturnDate, Now) > 0 Then
                        For I = 1 To 7
                            lst.ListSubItems.Item(I).ForeColor = &HFF&
                        Next I
                    End If
            .MoveNext
        Loop
    End If
    .Close
End With
End If
End Sub

Private Sub txtTotal_Change()
Call check
End Sub


Sub check()
If Me.txtTotal.Text = "0" And Me.ListView2.ListItems.count > 0 Then
    Me.bttnReturn.Enabled = True
Else
    Me.bttnReturn.Enabled = False
End If
End Sub


Private Sub bttnAdditional_Click()
 bReturn = True
    frmAddPenalty.Show vbModal
End Sub
