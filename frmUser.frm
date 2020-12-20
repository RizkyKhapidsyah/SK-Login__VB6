VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information Screen"
   ClientHeight    =   1848
   ClientLeft      =   2292
   ClientTop       =   3456
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1848
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   593
      ButtonWidth     =   487
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep0"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kClose"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kFinish"
            Object.ToolTipText     =   "Finish"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSave"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kDelete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kHelp"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdShowUsers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4080
      Picture         =   "frmUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Show Users"
      Top             =   600
      Width           =   372
   End
   Begin VB.CommandButton cmdPwdReset 
      Caption         =   "&Reset Password "
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   1560
   End
   Begin VB.ComboBox cboUserGroup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox txtExpirationDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM dd, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtUserID 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   684
      Width           =   2052
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtActivationDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "User's First Name"
      Top             =   1080
      Width           =   2532
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   3000
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
            Picture         =   "frmUser.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":09B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":0D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":105A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":13AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1702
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":20FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblUserID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label lblUserPassword 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   1332
   End
   Begin VB.Label lblUserExpireDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1080
      TabIndex        =   9
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Label lblUserActivationDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label lblUserTaskLevel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Security Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1572
   End
   Begin VB.Label lblUserFullName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1572
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'frmUsers - add/edit users
'
'******************************************************************
' Allow a user to be added, deleted, changes made, password re-sets
' Only allow administrators to access this area.
'******************************************************************
Dim mflag As Integer
Dim NewDate As Date
Dim userindex As Integer

Private Sub cboUserGroup_GotFocus()
    'hightlight character in cell
    
    cboUserGroup.SetFocus
    'start highlight before first character
    cboUserGroup.SelStart = 0
    'highlight to end of text
    cboUserGroup.SelLength = Len(cboUserGroup.Text)

End Sub

Private Sub cmdPwdReset_Click()
    RestUserPassword
    If mflag Then
        ClearAllFields
        txtPassword.Text = "password"
    End If
    HighLightText
End Sub
Private Sub cboUserGroup_Click()
    userindex = cboUserGroup.ListIndex
    'MsgBox cboUserGroup.List(userindex)
End Sub

Private Sub cmdShowUsers_Click()
    'open form to change password
    tShowuser = False
    frmShowUsers.Show vbModal
    If tShowuser Then
        txtUserID.Text = tUserID
        txtUserName.Text = tUserName
        cboUserGroup.Text = tUserGroup
        Me.Refresh
    End If
    txtUserID.SetFocus
    
End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    'add the task level to the combobox
    cboUserGroup.AddItem TASK_LEVEL_5
    cboUserGroup.AddItem TASK_LEVEL_4
    cboUserGroup.AddItem TASK_LEVEL_3
    cboUserGroup.AddItem TASK_LEVEL_2
    cboUserGroup.AddItem TASK_LEVEL_1
    cboUserGroup.AddItem TASK_LEVEL_0
    NewDate = Date
    txtActivationDate = Format(NewDate, "mm/dd/yyyy")
    NewDate = Date + NEXTMONTH
    txtExpirationDate = Format(NewDate, "mm/dd/yyyy")
    txtPassword.Text = Empty
    userindex = 0
    cboUserGroup.Text = ""
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kFinish"
            HandleFinishClick
        Case "kSave"
            HandleSaveClick
        Case "kCancel"
            'clear button
            HandleCancelClick
        Case "kDelete"
            HandleDeleteClick
        Case "kClose"
            HandleCloseClick
        Case "kHelp"
            HandleHelpClick
    End Select
End Sub

Public Sub AddEditUser()
    'add new user to table
    Dim UserID As String
    Dim Userpwd As String
    Dim UserName As String
    Dim Usergrp As String
    Dim UserActivationDate As Variant
    Dim UserExpirationDate As Variant
    Dim tableuserid As String
    Dim tablepwd As Variant
    Dim sSQL As String  'sql string
    
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    'return code
    Dim bRet As Boolean
    Dim EncryptData As String
    
    dospains.Initialize dsPA
    dospasel.Initialize dsPA

    sSQL = "select distinct userid, password from tblusers"
    dospasel.setSql sSQL
        
    If dospasel.execute(rdOpenStatic) = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "error_run1"
        'GoTo ERROR_run
    End If
    
    mflag = False
    UserID = Trim(txtUserID.Text)
    UserName = Trim(txtUserName.Text)
    Usergrp = cboUserGroup.Text  'cboUserGroup.List(userindex)
    'MsgBox cboUserGroup.List(userindex)
    Userpwd = Empty
    If UserID <> Empty And UserName <> Empty And Usergrp <> Empty Then
        'MsgBox "," & userid & "," & userpwd & ","
        'check for error, password must be new
        dospasel.moveFirst
        Do While Not dospasel.getEOF
            'check value to see if userid equals info in DB!
            tableuserid = Trim(dospasel.getRsValue(0))
            'tablepwd = dospasel.getRsValue(1)
            If UserID = tableuserid Then
                'user found
                mflag = True
                Exit Do
            End If
            dospasel.moveNext
        Loop
        If Not mflag Then
            UserActivationDate = CStr(Date)
            UserExpirationDate = CStr(Date + NEXTMONTH)
            sSQL = "insert into tblusers (userid, password, usergroup, username, activationdate, expirationdate)" _
                 + "values('" + UserID + "', '" + Userpwd + "', '" + Usergrp + "', '" + UserName + "', '" + UserActivationDate + "', '" + UserExpirationDate + "')"
            
            dospains.setSql sSQL
        
            If dospains.execute = False Then
                g_log.addEntry gc_iLogError, dospains.getError
                bRet = False
                'MsgBox "error_run1"
                GoTo ERROR_run
            End If
            g_log.addEntry gc_iLogInfo, "User Added to Database!"
            g_log.resetStat
        Else
            'userid already found, just update usergroup and username
            sSQL = "update tblusers set usergroup = '" + Usergrp + "'," _
                 + "username = '" + UserName + "'" _
                 + "where userid = '" + dospasel.getRsValue(0) + "' "
            
            dospains.setSql sSQL
        
            If dospains.execute = False Then
                g_log.addEntry gc_iLogError, dospains.getError
                bRet = False
                'MsgBox "error_run1"
                GoTo ERROR_run
            End If
            g_log.addEntry gc_iLogInfo, "User Updated in Database!"
            g_log.resetStat
        End If  'if not mflag, user not found
    Else
        'error empty fields
        Msg = "Please enter all information!"
        MsgBox Msg, 16, "Error!"
        'txtUserID.SetFocus
    End If  'info not correct
    Exit Sub
    
ERROR_run:
    ' show log
    If g_iRunMode <> 0 Then
        frmConsole.MousePointer = vbDefault
        Set frmLog.oLog = g_log
        frmLog.vLogStart = dtRunDate
        frmLog.Show vbModal
    End If
    g_log.resetStat

End Sub
Private Sub HandleFinishClick()
    'MsgBox "Finish.", vbInformation
    AddEditUser     'add or update user in table
    'clear fields on form
    ClearAllFields
    HighLightText
End Sub

Private Sub HandleSaveClick()
    'MsgBox "Saved.", vbInformation
    AddEditUser     'add or update user in table
    HighLightText
End Sub

Private Sub HandleCancelClick()
    'MsgBox "Cancel.", vbInformation
    ClearAllFields
    HighLightText
End Sub

Private Sub HandleDeleteClick() 'ByVal Button As ComctlLib.Button)
    'MsgBox "Deleted.", vbInformation
    DeleteUser
    If mflag Then
        ClearAllFields
    End If
    HighLightText
End Sub

Private Sub HandleCloseClick()
    'MsgBox "Close.", vbInformation
    Unload frmUser
End Sub

Private Sub HandleHelpClick()
    MsgBox "Display Help Screen.", vbInformation
    HighLightText
End Sub

Public Sub ClearAllFields()
    'clear all fields on form
    txtUserID.Text = ""
    txtUserName.Text = ""
    cboUserGroup.Text = ""
    NewDate = Date
    txtActivationDate = Format(NewDate, "mm/dd/yyyy")
    NewDate = Date + NEXTMONTH
    txtExpirationDate = Format(NewDate, "mm/dd/yyyy")
    txtPassword.Text = Empty
    userindex = 0
    Usergrp = ""
End Sub

Public Sub DeleteUser()
    'delete user from table
    Dim UserID As String
    Dim UserName As String
    Dim tableuserid As String
    Dim sSQL As String  'sql string
    
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    Dim dosPaDel As New CSQLDelete
    'return code
    Dim bRet As Boolean
    Dim EncryptData As String
    
    'MsgBox txtUserID.Text & "," & txtUserName.Text & "," & cboUserGroup.Text & ","
    
    dospains.Initialize dsPA
    dospasel.Initialize dsPA
    dosPaDel.Initialize dsPA

    sSQL = "select distinct userid from tblusers"
    dospasel.setSql sSQL
        
    If dospasel.execute(rdOpenStatic) = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "error_run1"
        'GoTo ERROR_run
    End If
    
    mflag = False
    UserID = Trim(txtUserID.Text)
    UserName = Trim(txtUserName.Text)
    If UserID <> "admin" Then
        'cannot delete user admin
        If UserID <> Empty And UserName <> Empty Then
            'MsgBox "," & userid & "," & userpwd & ","
            'check for error, password must be new
            dospasel.moveFirst
            Do While Not dospasel.getEOF
                'check value to see if userid equals info in DB!
                tableuserid = Trim(dospasel.getRsValue(0))
                If UserID = tableuserid Then
                    'user found
                    mflag = True
                    Exit Do
                End If
                dospasel.moveNext
            Loop
            If mflag Then
                If gUserID <> UserID Then
                    'user found, so delete user from table
                    sSQL = "delete from tblusers where userid = '" + dospasel.getRsValue(0) + "' "
                
                    dosPaDel.setSql sSQL
            
                    If dosPaDel.execute = False Then
                        g_log.addEntry gc_iLogError, dosPaDel.getError
                        bRet = False
                        'MsgBox "error_run1"
                        GoTo ERROR_run
                    End If
                Else
                    'cannot delete current user
                    Msg = "Cannot Delete Current User!"
                    MsgBox Msg, 16, "Error!"
                    mflag = False
                End If
            Else
                'error user not found
                Msg = "User not found!"
                MsgBox Msg, 16, "Error!"
                'txtUserID.SetFocus
            End If  'if not mflag, user not found
        Else
            'error empty fields
            Msg = "Please enter information!"
            MsgBox Msg, 16, "Error!"
            'txtUserID.SetFocus
        End If  'info not correct
    Else
        'cannot delete admin user
        Msg = "Cannot Delete Administrator - Admin!"
        MsgBox Msg, 16, "Error!"
        mflag = False
    End If  'userid = admin
    Exit Sub
    
ERROR_run:
    ' show log
    If g_iRunMode <> 0 Then
        frmConsole.MousePointer = vbDefault
        Set frmLog.oLog = g_log
        frmLog.vLogStart = dtRunDate
        frmLog.Show vbModal
    End If
    g_log.resetStat
End Sub

Public Sub HighLightText()
    'hightlight character in cell
    
    txtUserID.SetFocus
    'start highlight before first character
    txtUserID.SelStart = 0
    'highlight to end of text
    txtUserID.SelLength = Len(txtUserID.Text)

End Sub

Public Sub RestUserPassword()
    'rest the user password to empty
    Dim UserID As String
    Dim Userpwd As String
    Dim tableuserid As String
    Dim tablepwd As Variant
    Dim sSQL As String  'sql string
    Dim flag As Integer
    
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    'return code
    Dim bRet As Boolean
    Dim EncryptData As String
    
    dospains.Initialize dsPA
    dospasel.Initialize dsPA

    sSQL = "select distinct userid, password from tblusers"
    dospasel.setSql sSQL
        
    If dospasel.execute(rdOpenStatic) = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "error_run1"
        'GoTo ERROR_run
    End If
    
    mflag = False
    flag = True
    UserID = Trim(txtUserID.Text)
    Userpwd = Empty
    If UserID <> Empty Then
    
        If gUserID = UserID Then
            Msg = "When changing your password you will be required to log out before running any new tasks!"
            Style = vbYesNo + vbCritical + vbDefaultButton2
            Ans = MsgBox(Msg, Style, HM_MSG)
            'ans will return as IDYES = 6 or IDNO = 7
            flag = False
            If Ans = IDYES Then
                'unload the form
                flag = True
            End If
        End If

        If flag Then
            'MsgBox "," & userid & "," & userpwd & ","
            'check for error, password must be new
            dospasel.moveFirst
            Do While Not dospasel.getEOF
                'check value to see if userid equals info in DB!
                tableuserid = Trim(dospasel.getRsValue(0))
                'tablepwd = dospasel.getRsValue(1)
                If UserID = tableuserid Then
                    'user found
                    mflag = True
                    Exit Do
                End If
                dospasel.moveNext
            Loop
            If mflag Then
                'userid already found, just update Usergroup and username
                sSQL = "update tblusers set password = '" + Userpwd + "'" _
                     + "where userid = '" + dospasel.getRsValue(0) + "' "
            
                dospains.setSql sSQL
        
                If dospains.execute = False Then
                    g_log.addEntry gc_iLogError, dospains.getError
                    bRet = False
                    'MsgBox "error_run1"
                    GoTo ERROR_run
                End If
            Else
                'error user not found
                Msg = "User Not Found!"
                MsgBox Msg, 16, "Error!"
                mflag = False
            End If  'if not mflag, user not found
        End If  'if flag then run
    Else
        'error empty fields
        Msg = "Please enter information!"
        MsgBox Msg, 16, "Error!"
        'txtUserID.SetFocus
    End If  'info not correct
    Exit Sub
    
ERROR_run:
    ' show log
    If g_iRunMode <> 0 Then
        frmConsole.MousePointer = vbDefault
        Set frmLog.oLog = g_log
        frmLog.vLogStart = dtRunDate
        frmLog.Show vbModal
    End If
    g_log.resetStat
End Sub

Private Sub txtUserId_GotFocus()
    HighLightText
End Sub

Private Sub txtUserName_GotFocus()
    'hightlight character in cell
    
    txtUserName.SetFocus
    'start highlight before first character
    txtUserName.SelStart = 0
    'highlight to end of text
    txtUserName.SelLength = Len(txtUserName.Text)

End Sub
