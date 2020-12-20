VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter User ID and Password"
   ClientHeight    =   2220
   ClientLeft      =   3720
   ClientTop       =   4356
   ClientWidth     =   3984
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1306.525
   ScaleMode       =   0  'User
   ScaleWidth      =   3743.629
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox txtUID 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "If this is your first login, use 'Admin' for your User ID and 'Admin' for your Password."
      ForeColor       =   &H00C00000&
      Height          =   492
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3732
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "&Password:"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&User ID:"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''
'           Login form              '
'                                   '
'''''''''''''''''''''''''''''''''''''
'Option Explicit
'declare variables

Dim pFlag As Integer
Dim sConnect As String
Dim sDAOConnect As String
Dim sDsn As String
Dim strSQL As String

Private Sub cmdCancel_Click()
    'set the global var to false to denote a failed login
    badlogin = True
    End    'Unload frmLogin and end program
End Sub

Private Sub cmdOK_Click()
    
    Dim UserID As String
    Dim Userpwd As String
    Dim Usergroup As String
    Dim tableuserid As String
    Dim tablepwd As Variant
    Dim tablegrp As String
    Dim sSQL As String  'sql string
    
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    'return code
    Dim bRet As Boolean
    Dim flag As Integer
    Dim usrflag As Integer
    Dim EncryptData As String
    
    dospains.Initialize dsPA
    dospasel.Initialize dsPA
    
    'check for tblusers, if not found then create table
    sSQL = "select name from sysobjects"
    dospasel.setSql sSQL

    If dospasel.execute = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If

    'dosPasel.moveFirst
    bFound = False
    While Not dospasel.getEOF And bFound = False
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
        'MsgBox tmpstr & ","
        If tmpstr = "tblusers" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
            gFound = True
            'Exit Do
        End If
        dospasel.moveNext
    Wend
    If bFound = False Then
        'tblusers table was not found, create and populate the table
        g_oTable.run
    End If

    If gFound Then

        sSQL = "select distinct userid, password, usergroup, activationdate, expirationdate from tblusers order by userid"
    
        dospasel.setSql sSQL
        
        If dospasel.execute(rdOpenStatic) = False Then
            g_log.addEntry gc_iLogError, dospasel.getError
            bRet = False
            MsgBox "error_run1"
            'GoTo ERROR_run
        End If
    
        dospasel.moveFirst
        'check the IF statement, can't go over 3 tries!
        badlogin = True
        UserID = Trim(txtUID.Text)
        Userpwd = Trim(txtPWD.Text)
        'MsgBox "," & UserID & "," & Userpwd & ","
        flag = False
        usrflag = False
        'call pCounter to increment counter, 3x
        Call pCounter
        'check the IF statement, can't go over 3 tries!
        If pFlag <= 3 Then
            Do While Not dospasel.getEOF
                'MsgBox "," & UserID & "," & Userpwd & ","
                tableuserid = Trim(dospasel.getRsValue(0))
                tablepwd = Trim(dospasel.getRsValue(1))
                tablegrp = Trim(dospasel.getRsValue(2))
                'MsgBox "uid, " & UserID & " upwd, " & Userpwd & " tuid, " & tableuserid & " tpwd, " & tablepwd & ","
                'check value to see if userid equals info in DB!
                If UserID = tableuserid Then
                    Usergroup = tablegrp
                    'MsgBox "uid, " & UserID & " upwd, " & Userpwd & " tuid, " & tableuserid & " tpwd, " & tablepwd & ","
                    usrflag = True
                    If (tablepwd <> Empty) Then
                        DecryptData = Decrypt(tablepwd, EKey)
                        'MsgBox DecryptData & "," & tablepwd & ","         'Decrypt"
                        'MsgBox "," & tableuserid & "," & tablepwd & ","
                        If UserID = tableuserid And Userpwd = tablepwd Then
                            'set variable to TRUE, user found
                            flag = True
                            badlogin = False
                            Exit Do
                        End If
                    Else
                        'MsgBox "uid= " & UserID & " upwd= " & Userpwd & " tuid= " & tableuserid & " tpwd= " & tablepwd & ","
                        'MsgBox "need to ask for new password"
                        gUserID = UserID
                        gUserPwd = Userpwd
                        frmVerify.Show vbModal
                        If gFlag Then
                            flag = True
                            UserID = gUserID
                            Userpwd = gUserPwd
                            badlogin = False
                            Exit Do
                        Else
                            'error password cannot be null
                            badlogin = True
                            txtUID.SetFocus
                            Exit Sub
                        End If
                    End If  'not null
                End If  'userid is equal
                dospasel.moveNext
            Loop
            If Not flag Then
                'error user not found
                If usrflag Then
                    Msg = "Invalid Password, please re-enter the Password."
                Else
                    Msg = "Invalid User ID, please re-enter the User ID."
                End If
                MsgBox Msg, 16, "Login Error"
                badlogin = True
                txtUID.SetFocus
            Else
                gUserID = UserID
                gUserPwd = Userpwd
                gUserGroup = Usergroup
                If gLogin Then
                    'this form was opened from the frmcontrol form and
                    'not from the sub main, therefore, we need to re-open
                    'the frmcontrol.
                    gLogin = False
                    Unload frmLogin
                    frmControl.Show vbModal
                Else
                    Unload frmLogin
                End If
            End If
        Else
            'error user not found
            Msg = "Please Contact Administrator!"
            MsgBox Msg, 16, "Login Error"
            badlogin = True
            End
        End If  'if pflag
    Else
        Msg = "Please Contact Administrator!"
        MsgBox Msg, 16, "Login Error"
        badlogin = True
        End
    End If  'if gfound, table exist
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

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    pFlag = 0
End Sub

Private Sub txtPWD_GotFocus()
    'hightlight character in cell
    
    txtPWD.SetFocus
    'start highlight before first character
    txtPWD.SelStart = 0
    'highlight to end of text
    txtPWD.SelLength = Len(txtPWD.Text)

End Sub

Private Sub txtPWD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK.Value = True
End Sub

Private Sub txtUID_GotFocus()
    'hightlight character in cell
    
    txtUID.SetFocus
    'start highlight before first character
    txtUID.SelStart = 0
    'highlight to end of text
    txtUID.SelLength = Len(txtUID.Text)

End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If pFlag = 3 Then
        Call eEndprocess
    End If
    If KeyAscii = 13 Then KeyAscii = 0: txtPWD.SetFocus
    
End Sub

Public Sub pCounter()
    'use pFlag variable
    pFlag = pFlag + 1
End Sub

Public Sub eEndprocess()
    End 'end application
End Sub
