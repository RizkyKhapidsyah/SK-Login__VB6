VERSION 5.00
Begin VB.Form frmVerify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Required"
   ClientHeight    =   3084
   ClientLeft      =   3672
   ClientTop       =   3324
   ClientWidth     =   4608
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815.01
   ScaleMode       =   0  'User
   ScaleWidth      =   4329.98
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPWD2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   1932
   End
   Begin VB.TextBox txtPWD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   1932
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2400
      TabIndex        =   3
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   $"frmVerify.frx":0000
      Height          =   972
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   3972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm New Password:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password:"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1692
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''
' Verify Login form
'''''''''''''''''''''''''''''''''''''
'Option Explicit
'declare variables

Dim sConnect As String
Dim sDAOConnect As String
Dim sDsn As String
Dim strSQL As String

Private Sub cmdCancel_Click()
    'set the global var to false to denote a failed login
    badlogin = True
    Unload Me   'End program
    
End Sub

Private Sub cmdOK_Click()
    
    Dim Userpwd As String
    Dim Userpwd2 As String
    Dim tableuserid As String
    Dim tablepwd As Variant
    Dim sSQL As String  'sql string
    
    Dim dsPA As New CDsPA   'ds for selects only (from PA)
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    'return code
    Dim bRet As Boolean
    Dim flag As Integer
    Dim EncryptData As String
    
    dospains.Initialize dsPA
    dospasel.Initialize dsPA
    
    gFlag = False
    Userpwd = Trim(txtPWD.Text)
    Userpwd2 = Trim(txtPWD2.Text)
    If Userpwd = Userpwd2 And Userpwd <> Empty Then
        'MsgBox "passwords match"
        'but are they valid caharacters
        'check Userpwd for valid characters, between a and z or between A and Z
        flag = False
        ValidatePassword Userpwd, flag
        If flag Then
            sSQL = "select distinct userid, password from tblusers"
            dospasel.setSql sSQL
        
            If dospasel.execute(rdOpenStatic) = False Then
                g_log.addEntry gc_iLogError, dospasel.getError
                bRet = False
                'MsgBox "error_run1"
                'GoTo ERROR_run
            End If
    
            dospasel.moveFirst
            badlogin = True
            gFlag = False
            Do While Not dospasel.getEOF
                'check value to see if userid equals info in DB!
                tableuserid = Trim(dospasel.getRsValue(0))
                'tablepwd = Trim(dospasel.getRsValue(1))
                If tableuserid = gUserID Then
                    gUserPwd = Userpwd
                    EncryptData = Encrypt(Userpwd, EKey)
                    sSQL = "update tblusers set password = '" + Userpwd + "' where userid = '" + dospasel.getRsValue(0) + "' "
                    dospains.setSql sSQL
        
                    If dospains.execute = False Then
                        g_log.addEntry gc_iLogError, dospains.getError
                        bRet = False
                        'MsgBox "error_run1"
                        'GoTo ERROR_run
                    End If
                    gFlag = True
                    Exit Do
                End If
                dospasel.moveNext
            Loop
        Else
            Msg = "Invalid character in password!"
            MsgBox Msg, 16, "Error!"
        End If  'if flag
    Else
        If Userpwd <> Userpwd2 Then
            Msg = "Passwords do not match!"
            MsgBox Msg, 16, "Error!"
        Else
            If Userpwd = Empty Or Userpwd = Null Then
                Msg = "Invalid character in password!"
                MsgBox Msg, 16, "Error!"
            End If
        End If
        txtPWD.SetFocus
    End If  'password match
    If gFlag Then
        Unload frmVerify
    End If

End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    'display password in text box
    txtPWD.Text = gUserPwd
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

Private Sub txtPWD2_GotFocus()
    'hightlight character in cell
    
    txtPWD2.SetFocus
    'start highlight before first character
    txtPWD2.SelStart = 0
    'highlight to end of text
    txtPWD2.SelLength = Len(txtPWD2.Text)

End Sub

Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If pFlag = 3 Then
        Call eEndprocess
    End If
    If KeyAscii = 13 Then KeyAscii = 0: txtPWD.SetFocus
    
End Sub

Public Sub pCounter()
'use pFlag variable
    If txtUID.Text = "" Then
        txtUID.SetFocus
    Else
        pFlag = pFlag + 1
    End If

End Sub

Public Sub eEndprocess()
'end the program! To many tries.
    Unload frmLogin
End Sub

Public Sub ValidatePassword(Userpwd As String, flag As Integer)
    'check for invalid characters in the password
    'characters must alphanumeric, a to z, A to Z, or 0 to 9
    Dim K As Integer
    Dim Char As String
    Dim num As Integer
    Dim alphaflag As Integer
    Dim Lnum As Integer
    
    Lnum = Len(Userpwd)
    alphaflag = True
    For K = 1 To Lnum
        Char = Mid(Userpwd, K, 1)
        If Char = Empty Then
            alphaflag = False
            Exit For
        End If
        num = Asc(Char)
        'must be alphanumeric
        If (num >= 48 And num <= 57) Or (num >= 65 And num <= 90) Or (num >= 97 And num <= 122) Then
            'valid character
            'alphaflag = True
        Else
            'Msg = "Invalid character in password!"
            'MsgBox Msg, 64, "Error!"
            alphaflag = False
            Exit For
        End If
    Next K
    'if alphaflag then characters are valid
    If alphaflag Then
        flag = True
    End If

End Sub
