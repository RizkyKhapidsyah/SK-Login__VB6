VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Setup"
   ClientHeight    =   3264
   ClientLeft      =   2604
   ClientTop       =   2484
   ClientWidth     =   7236
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3264
   ScaleWidth      =   7236
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7236
      _ExtentX        =   12764
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep2"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kHelp"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLogFile 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modify the Windows registry"
      Height          =   1092
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   3492
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Items"
         Height          =   372
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   1092
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Items"
         Height          =   372
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1092
      End
   End
   Begin VB.TextBox txtPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtUserId 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "SQL Connection"
      Height          =   1332
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   3372
      Begin VB.Label lblDB 
         Alignment       =   1  'Right Justify
         Caption         =   "Database:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   852
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         Caption         =   "Server:"
         Height          =   252
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SQL Connection Authentication"
      Height          =   1332
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   3492
      Begin VB.Label lblPWD 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   852
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "Login ID:"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Log File Name and Location"
      Height          =   1092
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   3372
      Begin VB.Label lblLogFile 
         Alignment       =   1  'Right Justify
         Caption         =   "File Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   852
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4800
      Top             =   3960
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClear_Click()
    Msg = "This will clear the windows registry."
    Style = vbCritical + vbOK + vbDefaultButton2
    Ans = MsgBox(Msg, Style, "Setup")
    'ans will return as IDCANCEL = 2, IDOK = 1
    If Ans = IDOK Then
        g_parm.removeRAllParm
    End If
End Sub
Private Sub cmdUpdate_Click()
    Msg = "This will update the windows registry."
    Style = vbExclamation + vbOK + vbDefaultButton1
    Ans = MsgBox(Msg, Style, "Setup")
    'ans will return as IDCANCEL = 2, IDOK = 1
    If Ans = IDOK Then
        g_parm.setRParm gc_sSERVER, txtServer.Text
        g_parm.setRParm gc_sPADB, txtDB.Text
        g_parm.setRParm gc_sUID, txtUserID.Text
        g_parm.setRParm gc_sPWD, txtPWD.Text
        g_parm.setRParm gc_sLOG, txtLogFile.Text
        g_parm.setRParm gc_sCOMPLETE, "1"
        Unload frmSetup
    End If
End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    txtServer.Text = g_parm.getRParm(gc_sSERVER)
    txtDB.Text = g_parm.getRParm(gc_sPADB, "Northwind")
    txtUserID.Text = g_parm.getRParm(gc_sUID, "sa")
    txtPWD.Text = g_parm.getRParm(gc_sPWD)
    txtLogFile = g_parm.getRParm(gc_sLOG, "c:\mysql.log")
    
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    'kCancel is the key name of the button
    Select Case Button.Key
        Case "kCancel"
            HandleCancelClick
        Case "kHelp"
            HandleHelpClick
    End Select

End Sub
Private Sub HandleCancelClick()
    'MsgBox "Cancel.", vbInformation
    Unload frmSetup
End Sub

Private Sub HandleHelpClick()
    MsgBox "Display Help Screen.", vbInformation
End Sub

Private Sub txtDB_GotFocus()
    'hightlight character in cell
    
    txtDB.SetFocus
    'start highlight before first character
    txtDB.SelStart = 0
    'highlight to end of text
    txtDB.SelLength = Len(txtDB.Text)

End Sub

Private Sub txtLogFile_GotFocus()
    'hightlight character in cell
    
    txtLogFile.SetFocus
    'start highlight before first character
    txtLogFile.SelStart = 0
    'highlight to end of text
    txtLogFile.SelLength = Len(txtLogFile.Text)

End Sub

Private Sub txtPWD_GotFocus()
    'hightlight character in cell
    
    txtPWD.SetFocus
    'start highlight before first character
    txtPWD.SelStart = 0
    'highlight to end of text
    txtPWD.SelLength = Len(txtPWD.Text)

End Sub

Private Sub txtServer_GotFocus()
    'hightlight character in cell
    
    txtServer.SetFocus
    'start highlight before first character
    txtServer.SelStart = 0
    'highlight to end of text
    txtServer.SelLength = Len(txtServer.Text)

End Sub

Private Sub txtUserId_GotFocus()
    'hightlight character in cell
    
    txtUserID.SetFocus
    'start highlight before first character
    txtUserID.SelStart = 0
    'highlight to end of text
    txtUserID.SelLength = Len(txtUserID.Text)

End Sub
