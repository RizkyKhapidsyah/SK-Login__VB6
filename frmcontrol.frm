VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControl 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "SQL Database"
   ClientHeight    =   5808
   ClientLeft      =   2628
   ClientTop       =   2328
   ClientWidth     =   6540
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5808
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   5556
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5927
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "01/15/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:14 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbtoolbar 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kExit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep2"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kConsole"
            Object.ToolTipText     =   "Console"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep3"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCalc"
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep4"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCal"
            Object.ToolTipText     =   "Calendar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep5"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kHelp"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport crReport 
      Left            =   960
      Top             =   4800
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   120
      Top             =   4680
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":1DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":297C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":2D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcontrol.frx":2F00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangePass 
         Caption         =   "&Change User Password"
      End
      Begin VB.Menu mnuSpace0 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogin 
         Caption         =   "&User Login"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculater 
         Caption         =   "&Calculater"
      End
      Begin VB.Menu mnuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Ca&lendar"
      End
   End
   Begin VB.Menu mnuforms 
      Caption         =   "&Forms"
      Begin VB.Menu mnuconsole 
         Caption         =   "&Console"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuUserReport 
         Caption         =   "Database Users"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherReports 
         Caption         =   "Other Reports"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Change Database"
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddEditUsers 
         Caption         =   "Add/&Edit Users"
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogFile 
         Caption         =   "Log File"
         Begin VB.Menu mnuViewLog 
            Caption         =   "&View Log"
         End
         Begin VB.Menu mnuClearLog 
            Caption         =   "&Clear Log"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim flag As Integer

Private Sub Form_Load()
    'Center the form
    frmControl.Move (Screen.Width - frmControl.Width) / 2, (Screen.Height - frmControl.Height) / 2
    flag = True
    gLogin = False
    'need to check for administrator in usergroup,
    'if not admin then turn off setup options.
    'MsgBox gUserGroup
    If Trim(LCase(gUserGroup)) <> "administrator" Then
        mnuDatabase.Enabled = False
        mnuAddEditUsers.Enabled = False
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAddEditUsers_Click()
    'open form to change password
    frmUser.Show vbModal
End Sub

Private Sub mnuCalculater_Click()
    frmCalc.Show vbModal
End Sub

Private Sub mnuCalendar_Click()
    frmCalendar.Show vbModal
End Sub

Private Sub mnuChangePass_Click()
    'open form to change password
    frmChangePass.Show vbModal
End Sub

Private Sub mnuClearLog_Click()
    g_log.clearContents
    MsgBox "Log File has been cleared.", vbInformation
End Sub

Private Sub mnuconsole_Click()
    frmConsole.Show vbModal
End Sub

Private Sub mnuDatabase_Click()
    'setup database server one
    frmSetup.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Msg = "This will end your current session."
    Style = vbInformation + vbOK + vbDefaultButton1
    Ans = MsgBox(Msg, Style, "Session Exit")
    'ans will return as IDCANCEL = 2, IDOK = 1
    If Ans = IDOK Then
        End 'quit application
    End If
End Sub

Private Sub mnulogin_Click()
    'close control from and prompt for new user
    Msg = "This will end your current session."
    Style = vbInformation + vbOK + vbDefaultButton1
    Ans = MsgBox(Msg, Style, "Session Exit")
    'ans will return as IDCANCEL = 2, IDOK = 1
    If Ans = IDOK Then
        Unload frmControl
        gLogin = True
        frmLogin.Show vbModal
    End If
End Sub

Private Sub mnuOtherReports_Click()
    'can display other reports here
    MsgBox "Display Other Reports.", vbInformation
End Sub

Private Sub mnuUserReport_Click()
    crReport.ReportFileName = g_util.getExePath & "\users.rpt"
    crReport.Connect = _
        "uid=" & g_parm.getRParm(gc_sUID) & ";" _
        & "pwd=" & g_parm.getRParm(gc_sPWD) & ";" _
        & "driver={SQL Server};" _
        & "server=" & g_parm.getRParm(gc_sSERVER) & ";" _
        & "database=" & g_parm.getRParm(gc_sPADB)
    crReport.Action = 1

End Sub

Private Sub mnuViewLog_Click()
    Set frmLog.oLog = g_log
    frmLog.vLogStart = Null
    frmLog.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
    
End Sub

Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbtoolbar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbtoolbar.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    'when users click on the toolbar, run the code.
    'kExit is the key name of the button
    Select Case Button.Key
        Case "kExit"
            HandleExitClick
        Case "kCalc"            'calculator
            frmCalc.Show vbModal
        Case "kCal"             'calendar
            frmCalendar.Show vbModal
        Case "kConsole"
            frmConsole.Show vbModal
        Case "kHelp"
            HandleHelpClick
    End Select
End Sub

Private Sub HandleExitClick()
    'MsgBox "Exit Application.", vbInformation
    Msg = "This will end your current session."
    Style = vbInformation + vbOK + vbDefaultButton1
    Ans = MsgBox(Msg, Style, "Session Exit")
    'ans will return as IDCANCEL = 2, IDOK = 1
    If Ans = IDOK Then
        End 'quit application
    End If
End Sub

Private Sub HandleHelpClick()
    MsgBox "Display Help Screen.", vbInformation
End Sub
