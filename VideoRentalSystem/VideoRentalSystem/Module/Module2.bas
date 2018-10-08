Attribute VB_Name = "Module2"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public lst As ListViewItem

Public tType As String
Public oTampering As Boolean
Public oLose As Boolean
Public oDamage As Boolean
Public bPenalty As Boolean
Public bCDType As String
Public bReturn As Boolean


Public returnCount As Double
Public rentID As String
Public incomeID As String


Public dvdList As Boolean

Public logoutDate As String
Public loginTime As Date
Public logOutTime As String
Public user As String
Public ctr As Double

Public OK As Boolean
Public FIRST As Boolean
Public xEdit As Boolean
Public xUsr As String



Public wDone As String


Sub Main()
Set cn = New ADODB.Connection
    With cn
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=VideoRentalsSystem"
        .Open
    End With
        Call checkUser
End Sub

Public Sub checkUser()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Users", cn, 1, 2
    
    If .EOF = False Then
        frmLogin.Show
    Else
        MsgBox "This is the first run of the system. Please register the Administrator", vbInformation + vbOKOnly
        FIRST = True
        frmAddUser.Show vbModal
    End If
    .Close
End With
End Sub
Public Sub auditT()
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_AuditTrail", cn, 1, 2
    
    .AddNew
    !UserID = user
    !Date = Now
    !wDone = wDone
    .Update
    
    .Close
End With
End Sub


Sub out()
    logoutDate = Date
    logOutTime = Time
    
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_LogHistory where logintime like '" & loginTime & "'", cn, 1, 2
        
        !logOutTime = logOutTime
        !logoutDate = logoutDate
        .Update
        
    End With
End Sub

Public Sub chckMember(ByVal ID As String)
    Dim ctr As Integer
    ctr = 0
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from tbl_DVDRented where Status like '" & "RENTED" & "'", cn, 1, 2
    
        Do Until .EOF
        If ID = !MemberID Then
            ctr = ctr + 1
            .MoveNext
        Else
            .MoveNext
        End If
    Loop
    .Close
    End With

    If ctr = 3 Then
        OK = False
    Else
        OK = True
    End If
End Sub

Public Sub getIncomeID()
incomeID = 1
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from tbl_Income", cn, 1, 2
    
    If .EOF = True Then
        incomeID = 1
    Else
        Do Until .EOF
            incomeID = incomeID + 1
            .MoveNext
        Loop
    End If
    .Close
End With
End Sub


Function Text_Empty(ByVal oText As FlatEdit) As Boolean
If Trim(oText.Text) = "" Or Trim(oText.Text) = "Last Name" Or Trim(oText.Text) = "First Name" Or Trim(oText.Text) = "Middle Name" Then
    oText.BackColor = &HC0C0C0
    oText.SetFocus
    Text_Empty = True
Else
    oText.BackColor = &HFFFFFF
    Text_Empty = False
End If
End Function
Function Combo_Empty(ByVal oText As ComboBox) As Boolean
If Trim(oText.Text) = "" Then
    oText.BackColor = &HC0C0C0
    oText.SetFocus
    Combo_Empty = True
Else
    oText.BackColor = &HFFFFFF
    Combo_Empty = False
End If
End Function


