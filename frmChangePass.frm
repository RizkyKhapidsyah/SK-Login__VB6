VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1728
   ClientLeft      =   3372
   ClientTop       =   3108
   ClientWidth     =   4356
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1728
   ScaleWidth      =   4356
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   4356
      _ExtentX        =   7684
      _ExtentY        =   593
      ButtonWidth     =   487
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kFinish"
            Object.ToolTipText     =   "Finish"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtActivationDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1365
   End
   Begin VB.TextBox txtNewPwd2 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   2160
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2160
   End
   Begin VB.TextBox txtOldPwd 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2880
      Width           =   2520
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1392
   End
   Begin VB.TextBox txtExpirationDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1404
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   2148
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   3360
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
            Picture         =   "frmChangePass.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChangePass.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ActivationDate:"
      Height          =   288
      Index           =   3
      Left            =   612
      TabIndex        =   13
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Password Verify:"
      Height          =   288
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New password:"
      Height          =   252
      Left            =   36
      TabIndex        =   11
      Top             =   960
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      Height          =   252
      Left            =   36
      TabIndex        =   10
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   288
      Index           =   2
      Left            =   744
      TabIndex        =   9
      Top             =   3252
      Width           =   1284
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ExpirationDate:"
      Height          =   288
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   3948
      Width           =   1308
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserID:"
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
' frmChangePass - used for changing the password
'
'***********************************************
' This code will allow you to change password.
' When changing passwords a user can not select
' the DEFAULT_PASSWORD or any string that contains
' "password", this is very common among users.
'************************************************
Dim mflag As Integer

Sub Change_Pass(mflag As Integer)
    Dim UserID As String
    Dim Userpwd As String
    Dim Userpwd2 As String
    Dim UserNewpwd As String
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
    
    pFlag = 0
    badlogin = True
    UserID = gUserID
    Userpwd = Trim(txtOldPwd.Text)
    UserNewpwd = Trim(txtNewPwd.Text)
    'MsgBox "," & userid & "," & userpwd & ","
    'check for error, password must be new
    dospasel.moveFirst
    Do While Not dospasel.getEOF
        'check value to see if userid equals info in DB!
        tableuserid = Trim(dospasel.getRsValue(0))
        tablepwd = dospasel.getRsValue(1)
        If tablepwd <> Empty Or tablepwd <> Null Then
            tablepwd = Trim(dospasel.getRsValue(1))
            DecryptData = Decrypt(tablepwd, EKey)
            'MsgBox "userid= " & UserID & ", tableid= " & tableuserid & ", userpwd= " & Userpwd & ", tablepwd= " & tablepwd
            If UserID = tableuserid And Userpwd = tablepwd Then
                'set variable to TRUE, user found
                mflag = True
                badlogin = False
                gUserPwd = UserNewpwd       'Update Global password
                EncryptData = Encrypt(UserNewpwd, EKey)
                'MsgBox "Encry= " & EncryptData & ", Newpwd= " & UserNewpwd & ","         'EKey
                txtActivationDate.Text = Date
                txtExpirationDate.Text = Date + NEXTMONTH
                gActivationDate = txtActivationDate.Text
                gExpirationDate = txtExpirationDate.Text
                'sSQL = "update tblusers set password = '" + UserNewpwd + "' where userid = '" + dospasel.getRsValue(0) + "' "
                sSQL = "update tblusers set password = '" + UserNewpwd + "'," _
                     + "activationdate = '" + gActivationDate + "'," _
                     + "expirationdate = '" + gExpirationDate + "'" _
                     + "where userid = '" + dospasel.getRsValue(0) + "' "
                dospains.setSql sSQL
        
                If dospains.execute = False Then
                    g_log.addEntry gc_iLogError, dospains.getError
                    bRet = False
                    'MsgBox "error_run1"
                    'GoTo ERROR_run
                End If
                Exit Do
            End If  'if userid and userpwd found
        End If  'if not null
        dospasel.moveNext
    Loop
    If Not mflag Then
        'error user not found
        Msg = "Password not changed!"
        MsgBox Msg, 16, "Error!"
        badlogin = True
        txtNewPwd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    txtUserID.Text = gUserID
    'get dates from table
    GetDates
    txtPassword.Text = gUserPwd
    txtOldPwd.Text = gUserPwd
End Sub

Private Sub txtNewPwd_GotFocus()
    'hightlight character in cell
    txtNewPwd.SetFocus
    'start highlight before first character
    txtNewPwd.SelStart = 0
    'highlight to end of text
    txtNewPwd.SelLength = Len(txtNewPwd.Text)

End Sub
Private Sub txtNewPwd_Validate(KeepFocus As Boolean)
    Dim count As String
    Dim SearchStr, PassStr, MyStr As String
    
    count = Len(txtNewPwd.Text)
    'Make sure password does not equal old password
    If txtNewPwd.Text = txtOldPwd.Text Then
        KeepFocus = True
        MsgBox "Your password can be the same as the old.", vbApplicationModal, "New Password"
        Exit Sub
    End If
    'Now make sure it is long enough, or they will use ABC then 123
    If count < MINIMUM_PASSWORD_LENGTH And txtNewPwd.Text <> "" Then
        KeepFocus = True
        MsgBox "Your password must be at least 5 characters.", vbApplicationModal, "New Password"
        txtNewPwd.Text = ""
        Exit Sub
    End If
    'Make sure password does not equal password
    If txtNewPwd.Text = DEFAULT_PASSWORD Then
        KeepFocus = True
        MsgBox "The default 'password' can not be used.", vbCritical, "New Password."
        txtNewPwd.Text = ""
        Exit Sub
    End If
    'Make sure password is not in the string, they may use password1, 2,3...
    PassStr = "password"
    SearchStr = txtNewPwd.Text
    MyStr = InStr(1, SearchStr, PassStr, vbTextCompare)
    If MyStr <> 0 Then
        KeepFocus = True
        MsgBox "The default 'password' can not be in the password.", vbCritical, "New Password."
        txtNewPwd.Text = ""
        Exit Sub
    End If
    'Make sure the password is not the user id
    If txtNewPwd.Text = gUserID Then
        KeepFocus = True
        MsgBox "The 'password' can not be the UserID.", vbCritical, "New Password."
        txtNewPwd.Text = ""
        Exit Sub
    End If
    'this checks the new password against the old
    'to make sure they are not similiar words
    Dim midstr, midpassstr, newmidstr As String
    midstr = txtOldPwd.Text
    midpassstr = midstr
    newmidstr = Mid(midstr, 3, 4)
    SearchStr = txtNewPwd.Text
    MyStr = InStr(1, SearchStr, newmidstr, vbTextCompare)
    If MyStr <> 0 Then
        KeepFocus = True
        MsgBox "The password can not contain similiar words from old password.", vbCritical, "New Password."
        txtNewPwd.Text = ""
        Exit Sub
    End If
End Sub
Private Sub txtNewPwd2_Change()
    Dim count As String
    Dim count1 As String
    
    count = Len(txtNewPwd.Text)
    count1 = Len(txtNewPwd2.Text)
End Sub

Private Sub txtNewPwd2_GotFocus()
    'hightlight character in cell
    
    txtNewPwd2.SetFocus
    'start highlight before first character
    txtNewPwd2.SelStart = 0
    'highlight to end of text
    txtNewPwd2.SelLength = Len(txtNewPwd2.Text)

End Sub

'Verify that the two new passwords are equal
Private Sub txtNewPwd2_Validate(KeepFocus As Boolean)
    Dim flag As Integer
    
    flag = False
    If txtNewPwd2.Text <> txtNewPwd.Text And txtNewPwd2.Text <> "" Then
        MsgBox "Your password and verify password do not match.", vbApplicationModal, "Password"
        txtNewPwd.Text = ""
        txtNewPwd2.Text = ""
        txtNewPwd.SetFocus
        flag = True
    End If
End Sub

Public Sub GetDates()
    'get dates from table
    Dim UserID As String
    Dim tableuserid As String
    Dim sSQL As String  'sql string
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    
    dospasel.Initialize dsPA

    sSQL = "select distinct userid, password, activationdate, expirationdate from tblusers"
    dospasel.setSql sSQL
        
    If dospasel.execute(rdOpenStatic) = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "error_run1"
        'GoTo ERROR_run
    End If
    
    UserID = gUserID
    dospasel.moveFirst
    Do While Not dospasel.getEOF
        'check value to see if userid equals info in DB!
        tableuserid = Trim(dospasel.getRsValue(0))
        If UserID = tableuserid Then
            txtActivationDate.Text = dospasel.getRsValue(2)
            txtExpirationDate.Text = dospasel.getRsValue(3)
            gActivationDate = txtActivationDate.Text
            gExpirationDate = txtExpirationDate.Text
            Exit Do
        End If  'if userid and userpwd found
        dospasel.moveNext
    Loop
End Sub
Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kFinish"
            HandleFinishClick
        Case "kCancel"
            'clear button
            HandleCancelClick
    End Select

End Sub
Private Sub HandleFinishClick()
    'MsgBox "Finish.", vbInformation
    mflag = False
    If txtNewPwd.Text <> "" And txtNewPwd2.Text <> "" Then
        Call Change_Pass(mflag)
        If mflag Then
            Unload frmChangePass
        End If
    Else
        'entries cannot be empty
        Msg = "Please Enter Password."
        MsgBox Msg, 16, "Error!"
        badlogin = True
        txtNewPwd.SetFocus
    End If
End Sub

Private Sub HandleCancelClick()
    Unload frmChangePass
End Sub



