VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.0#0"; "COC8F1~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.0#0"; "CODEJO~3.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.0#0"; "COA766~1.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Video Rental System"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   0
      Top             =   0
      _Version        =   983040
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   2880
      Top             =   4800
      _Version        =   983040
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   3
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3720
      Top             =   4320
      _Version        =   983040
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   3
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Dim themes As String

Dim WithEvents Workspace As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1
Dim WorkspaceVisible As Boolean

Dim I As Integer
Const IMAGEBASE = 10000

Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
End Function



Private Sub MDIForm_Initialize()
InitCommonControls
End Sub

Private Sub Workspace_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    Debug.Print "Workspace_SelectedChanged. Item.Caption = " & Item.Caption
End Sub

Public Sub SetCaption(Caption As String)
    SetWindowText Me.hWnd, Caption
End Sub

Private Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, ID, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub CommandBars_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    Options.ShowRibbonQuickAccessPage = True
    
    Dim Controls As CommandBarControls
    Set Controls = CommandBars.DesignerControls
    
    If (Controls.count = 0) Then
        AddButton Controls, xtpControlButton, ID_PRODUCT_ENTRY, "&New CD/DVD Entry", False, "Create CD/DVD", xtpButtonAutomatic, "File"
        AddButton Controls, xtpControlButton, ID_PRODUCT_RENT, "Rent &CD/DVD", False, "Rent CD/DVD", xtpButtonAutomatic, "File"
        AddButton Controls, xtpControlButton, ID_PRODUCT_RETURN, "&Return CD/DVD", False, "Return CD/DVD", xtpButtonAutomatic, "File"
        
        
        
        
'        AddButton Controls, xtpControlButton, ID_MAINTENANCE_USER_LIST, "&User Lists", False, "Create/modify/delete User settings", xtpButtonAutomatic, "Maintenance"
'
'
'        AddButton Controls, xtpControlButton, ID_MAINTENANCE_DISCPENALTIES, "&Disc Penalties", False, "Insert Clipboard contents", xtpButtonAutomatic, "Maintenance"
              
        AddButton Controls, xtpControlButton, ID_APP_ABOUT, "About", False, "", xtpButtonAutomatic, "Help"
    End If
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.ID
        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
            Debug.Print "Selected Tab is Changed"
        
        Case XTP_ID_RIBBONCUSTOMIZE:
            CommandBars.ShowCustomizeDialog 3
        Case ID_APP_ABOUT:
            MsgBox "Programmers:"
            
            
            
        Case ID_PRODUCT_ENTRY:
            frmNewCDEntry.Show vbModal
            
        Case ID_TRANSACTION_RENT:
            frmRent.Show
         Case ID_TRANSACTION_RETURN:
            frmReturn.Show
             Case ID_PRODUCT_RENT:
        frmRent.Show
        Case ID_PRODUCT_ADDEXISTING:
            frmAddExisting.Show vbModal
            Case ID_PRODUCT_NEWMEMBER:
           ' frmNewMember.Show
        
            

        Case ID_MAINTENANCE_DISCRENTPERIOD:
            frmRentPeriod.Show
        Case ID_MAINTENANCE_DISCPENALTIES:
            frmAddPenalty.Show
        Case ID_MAINTENANCE_DISCRENTALRATE:
            frmRentalRate.Show
         Case ID_MAINTENANCE_USER_CHANGEPASS:
          frmChangePass.Show vbModal
               Case ID_MAINTENANCE_MODIFYMEMBERS:
            frmMembers.Show
           Case ID_MAINTENANCE_MEMBERPAYMENT
            frmMembershipPaymentSettings.Show
        
        
        Case ID_INVENTORY_LISTCD:
             frmMemberList.Show
        Case ID_INVENTORY_LISTMEMBERS:
            frmDVDList.Show
        Case ID_INVENTORY_UNRETURNCD:
            frmUnreturn.Show
    
         
        Case ID_PRODUCT_RETURN:
                frmReturn.Show
        Case ID_REPORT_LOGHISTORY:
                frmLogHistory.Show
                  
        
            
            Case ID_TOOLS_NOTEPAD:
             Shell "Notepad.exe", vbNormalFocus
             Case ID_TOOLS_CALCULATOR:
            Shell "Calc.exe", vbNormalFocus
            
            Case ID_WINDOWSETTINGS_THEMES:
            frmThemes.Show vbModal
            
            Case ID_SYSTEMLOCK_LOCK:
                frmLockSystem.Show
        Case ID_APP_LOGOUT:
        
            Call logout


                                            
        Case Else:
            MsgBox Control.Caption & " clicked", vbOKOnly, "Button Clicked"
    End Select
End Sub

Private Sub LoadIcons()
    CommandBars.Options.UseSharedImageList = False

    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Media-CD-RW.ico", ID_PRODUCT_ENTRY, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\cd-burner-copy.ico", ID_PRODUCT_RENT, XtremeCommandBars.XTPImageState.xtpImageNormal
   CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Download.ico", ID_PRODUCT_RETURN, XtremeCommandBars.XTPImageState.xtpImageNormal


    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Stocks.ico", ID_MAINTENANCE_DISCPENALTIES, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\cashbox.ico", ID_MAINTENANCE_DISCRENTALRATE, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Toolbox.ico", ID_MAINTENANCE_DISCRENTPERIOD, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\User-Group.ico", ID_MAINTENANCE_USER_LIST, XtremeCommandBars.XTPImageState.xtpImageNormal

    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Apps-preferences-desktop-user.ico", ID_MAINTENANCE_MODIFYMEMBERS, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\skills.ico", ID_MAINTENANCE_MEMBERPAYMENT, XtremeCommandBars.XTPImageState.xtpImageNormal



    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\CD-Drive.ico", ID_PRODUCT_ADDEXISTING, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Actions-user-group-new.ico", ID_PRODUCT_NEWMEMBER, XtremeCommandBars.XTPImageState.xtpImageNormal


    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\cd-burner-copy.ico", ID_TRANSACTION_RENT, XtremeCommandBars.XTPImageState.xtpImageNormal
   CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Download.ico", ID_TRANSACTION_RETURN, XtremeCommandBars.XTPImageState.xtpImageNormal
   
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\notepad(1).ico", ID_INVENTORY_LISTCD, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Hard-Drive.ico", ID_INVENTORY_LISTMEMBERS, XtremeCommandBars.XTPImageState.xtpImageNormal
     CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Hard-Drive.ico", ID_INVENTORY_UNRETURNCD, XtremeCommandBars.XTPImageState.xtpImageNormal

    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Paper-Money.ico", ID_REPORT_INCOME, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Folder-User(1).ico", ID_REPORT_AUDITTRAIL, XtremeCommandBars.XTPImageState.xtpImageNormal
    
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Actions-view-history.ico", ID_REPORT_LOGHISTORY, XtremeCommandBars.XTPImageState.xtpImageNormal


    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Notepad.ico", ID_TOOLS_NOTEPAD, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\calculator.ico", ID_TOOLS_CALCULATOR, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\archive.ico", ID_TOOLS_ARCHIVE, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Misc-Download-Database.ico", ID_TOOLS_BACKUP, XtremeCommandBars.XTPImageState.xtpImageNormal


    CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Configuration-Settings.ico", ID_MAINTENANCE_USER_CHANGEPASS, XtremeCommandBars.XTPImageState.xtpImageNormal
 


   CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\Categories-preferences-other.ico", ID_WINDOWSETTINGS_THEMES, XtremeCommandBars.XTPImageState.xtpImageNormal
'
  CommandBars.Icons.LoadIcon App.Path & "\UsedIcons\unlock-server.ico", ID_SYSTEMLOCK_LOCK, XtremeCommandBars.XTPImageState.xtpImageNormal


    CommandBars.Icons.LoadBitmap App.Path & "\dialog_information.png", ID_APP_ABOUT, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadIcon App.Path & "\res\GroupPopup.ico", ID_GROUP_POPUPICON, XtremeCommandBars.XTPImageState.xtpImageNormal


    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext

    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    ToolTipContext.ShowShadow = True
End Sub




Private Sub CreateRibbonBar()
    Dim ControlUser As CommandBarControl
    
    Dim TabReport As RibbonTab
    Dim TabTransaction As RibbonTab
    Dim TabMaintenance As RibbonTab
    Dim TabInventory As RibbonTab
    Dim TabTools As RibbonTab
    Dim TabWindowSettings As RibbonTab
    Dim TabSystemLock As RibbonTab
    
    
  
    Dim GroupFile As RibbonGroup
    Dim GroupMaintenance As RibbonGroup
    Dim GroupInventory As RibbonGroup
    Dim GroupReport As RibbonGroup
    Dim GroupTools As RibbonGroup
    Dim GroupWindowSettings As RibbonGroup
    Dim GroupSystemLock As RibbonGroup
    

    Dim ControlCategory As CommandBarPopup
    Dim Control As CommandBarControl

    Dim ControlPopup As CommandBarPopup
    Dim ControlMargins As CommandBarPopup
    Dim ControlOrientation As CommandBarPopup
    Dim ControlSize As CommandBarPopup
    Dim ControlFile As CommandBarPopup
    Dim ControlAbout As CommandBarControl

    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    'RibbonBar.Customizable = True

    Set ControlFile = RibbonBar.AddSystemButton()
    ControlFile.IconId = ID_SYSTEM_ICON
    ControlFile.CommandBar.Controls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_ENTRY, "&New CD/DVD Entry", False, False
    ControlFile.CommandBar.Controls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_RENT, "&Rent CD/DVD", False, False
    ControlFile.CommandBar.Controls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_RETURN, "&Return CD/DVD", False, False
    
    Set Control = ControlFile.CommandBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlButton, ID_APP_LOGOUT, "&Logout", False, False)
    Control.BeginGroup = True
    ControlFile.CommandBar.SetIconSize 32, 32

    Set ControlAbout = RibbonBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlButton, ID_APP_ABOUT, "&About", False, False)
    ControlAbout.Flags = XtremeCommandBars.XTPControlFlags.xtpFlagRightAlign

    Set TabTransaction = RibbonBar.InsertTab(0, "&Transaction")
    TabTransaction.ID = ID_TAB_TRANSACTION

    Set GroupFile = TabTransaction.Groups.AddGroup("Transaction", ID_GROUP_TRANSACTION)
    GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TRANSACTION_RENT, "Rent &CD/DVD", False, False
    GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TRANSACTION_RETURN, "&Return CD/DVD", False, False
   
    
    Set GroupFile = TabTransaction.Groups.AddGroup("CD/ DVD Entry", ID_GROUP_TRANSACTION)
    GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_ENTRY, "&New CD/DVD Entry", False, False
    GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_ADDEXISTING, "&Add Existing CD/DVD", False, False
   
    
    Set GroupFile = TabTransaction.Groups.AddGroup("Member Entry", ID_GROUP_TRANSACTION)
     GroupFile.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_NEWMEMBER, "N&ew Member ", False, False
    
    
    Set TabMaintenance = RibbonBar.InsertTab(1, "&Maintenance")
    TabMaintenance.ID = ID_TAB_MAINTENANCE

    Set GroupMaintenance = TabMaintenance.Groups.AddGroup("User Settings", ID_GROUP_MAINTENANCE)
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_USER_CHANGEPASS, "&Change Password", False, False
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_USER_LIST, "&User Lists", False, False
     
     
    Set GroupMaintenance = TabMaintenance.Groups.AddGroup("Maintenance", ID_GROUP_MAINTENANCE)
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_DISCPENALTIES, "&Disk Penalty ", False, False
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_DISCRENTALRATE, "D&isk Rental Rate", False, False
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_DISCRENTPERIOD, "Di&sk Rent Period", False, False
    
    Set GroupMaintenance = TabMaintenance.Groups.AddGroup("Member Settings", ID_GROUP_MAINTENANCE)
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_MODIFYMEMBERS, "&Modify Members", False, False
    GroupMaintenance.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_MAINTENANCE_MEMBERPAYMENT, "Membership &Payment ", False, False
    

    Set TabReport = RibbonBar.InsertTab(3, "&Reports")
    TabReport.ID = ID_TAB_REPORT
    Set GroupReport = TabReport.Groups.AddGroup("&Reports", ID_GROUP_REPORT)
    GroupReport.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_REPORT_INCOME, "&Income Report ", False, False
    GroupReport.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_REPORT_AUDITTRAIL, "&Audit Trail", False, False
    GroupReport.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_REPORT_LOGHISTORY, "&Log History", False, False
    
    
    Set TabTools = RibbonBar.InsertTab(4, "&Tools")
    TabTools.ID = ID_TAB_TOOLS
    Set GroupTools = TabTools.Groups.AddGroup("&Tools", ID_GROUP_TOOLS)
    GroupTools.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TOOLS_NOTEPAD, "&Notepad", False, False
    GroupTools.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TOOLS_CALCULATOR, "&Calculator", False, False
    GroupTools.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TOOLS_ARCHIVE, "&Archive", False, False
    GroupTools.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_TOOLS_BACKUP, "&Back-up Database", False, False
    
    Set TabWindowSettings = RibbonBar.InsertTab(5, "&Window Settings")
    TabWindowSettings.ID = ID_TAB_WINDOWSETTINGS
    Set GroupWindowSettings = TabWindowSettings.Groups.AddGroup("&Themes", ID_GROUP_WINDOWSETTINGS)
    GroupWindowSettings.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_WINDOWSETTINGS_THEMES, "&Change Themes", False, False
    
   
    Set TabSystemLock = RibbonBar.InsertTab(6, "&System Lock")
    TabSystemLock.ID = ID_TAB_SYSTEMLOCK
    Set GroupSystemLock = TabSystemLock.Groups.AddGroup("&Lock", ID_GROUP_WINDOWSETTINGS)
    GroupSystemLock.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_SYSTEMLOCK_LOCK, "&Lock", False, False
     
    Control.Checked = True
   
    Set TabInventory = RibbonBar.InsertTab(2, "&Drawer")
    TabInventory.ID = ID_TAB_UTILITIES
    Set GroupInventory = TabInventory.Groups.AddGroup("Record", ID_GROUP_DRAWER)
    GroupInventory.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_INVENTORY_LISTCD, "List of &Members", False, False
     GroupInventory.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_INVENTORY_LISTMEMBERS, "&List of CDs/DVDs", False, False
    GroupInventory.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_INVENTORY_UNRETURNCD, "List of &Unreturn CDs/DVDs", False, False
   
    
    RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_ENTRY, "&New CD/DVD Entry", False, False
    RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_RENT, "&Rent CD/DVD", False, False
    RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_PRODUCT_RETURN, "&Return CD/DVD", False, False
End Sub

Private Sub MDIForm_Load()

   ' For I = 0 To 2
   ' frmThemes.rBttn(I).value = False
   ' Next I

    Dim Descriptions As SkinDescriptions
    Set Descriptions = frmThemes.SkinFramework.EnumerateSkinDirectory(App.Path + "..\..\VideoRentalSystem\Styles", True)
    
    Dim S As SkinDescription
    For Each S In Descriptions
        Debug.Print S.Name & " - " & S.Path
    Next
  
   ' Set rs = New ADODB.Recordset
    'With rs
    '.Open "Select * from tbl_Themes", cn, 1, 2

    'frmThemes.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\" + !themes, ""
    'frmThemes.SkinFramework.ApplyWindow Me.hWnd
    'frmThemes.SkinFramework.ApplyOptions = frmThemes.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
   'End With
    
     MDIMain.SkinFramework.LoadSkin App.Path + "..\..\VideoRentalSystem\Styles\Vista.cjstyles", ""
        MDIMain.SkinFramework.ApplyWindow Me.hWnd
        MDIMain.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

    LoadIcons
    CreateRibbonBar
    
    Dim StatusBar As StatusBar
    Set StatusBar = CommandBars.StatusBar
    StatusBar.Visible = True
    
    StatusBar.AddPane 0

    
'    RibbonBar.EnableFrameTheme
    
    CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowWindowsDefault
        
    CommandBars.EnableCustomization True
    
    WorkspaceVisible = True
    CommandBars.ShowTabWorkspace True
    
    DockingPaneManager.SetCommandBars Me.CommandBars

End Sub

Private Function GetMDIChildCount() As Boolean
   GetMDIChildCount = IIf(Not ActiveForm Is Nothing, True, False)
End Function
  
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Do you want to exit?", vbCritical + vbYesNo) = vbYes Then
    Call out
    End
    UnloadMode = 1
    Cancel = 0
End If
    Cancel = 1
    UnloadMode = 0
End Sub

Private Sub logout()

On Error Resume Next
If MsgBox("Are you want to log out?", vbQuestion + vbYesNo) = vbYes Then
    Call out
    Unload MDIMain.ActiveForm
    MDIMain.Hide
    
    MsgBox "Successfuly logout.", vbInformation + vbOKOnly
    frmLogin.Show
End If
End Sub





