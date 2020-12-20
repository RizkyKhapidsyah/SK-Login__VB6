Attribute VB_Name = "basUserInformation"
Option Explicit
'*******************************************************
'UserInformation Module
'
'Copyright: (c) 2001 MicroNet Software Technology
'
'*******************************************************
Global gUserID As String
Global gUserPwd As String
Global gUpdateUser As String
Global gUpdatePwd As String
Global gFlag As Integer     'verify password was changed
Global gUserName As String
Global gExpirationDate As String
Global gActivationDate As String
Global gUserGroup As String
Global gLogin As Integer
Global gFound As Integer
Global tUserID As String
Global tUserName As String
Global tUserGroup As String
Global tShowuser As Integer
'User specific constants
Global Const EXPIRE_TERM = 180 'password expiration interval in days
Global Const MINIMUM_PASSWORD_LENGTH = 5 'Minimum password length
Global Const DEFAULT_PASSWORD = "password"
Global Const NEXTYEAR = 365     '365 days - password expires
Global Const NEXTMONTH = 30     '30 days - password expires
Global Const TASK_LEVEL_5 = "Administrator" 'You can name the levels what ever you want
Global Const TASK_LEVEL_4 = "Billing"
Global Const TASK_LEVEL_3 = "Payroll"
Global Const TASK_LEVEL_2 = "Accountant"
Global Const TASK_LEVEL_1 = "User"
Global Const TASK_LEVEL_0 = ""

'encryption key
Global Const EKey = "94022" 'Chr$(57) & Chr$(52) & Chr$(48) & Chr$(50) & Chr$(50)  'encrypt key - "94022"

Public Function Encrypt(Secret As Variant, CryptKey As Variant) As String
      'secret = the string you wish to encrypt or decrypt.
      'CryptKey = the password (EKey=94022) used to encrypt the string.
      Dim L%, X%, Char%

      L = Len(CryptKey)
      For X = 1 To Len(Secret)
         Char = Asc(Mid(CryptKey, (X Mod L) - L * ((X Mod L) = 0), 1))
         Mid(Secret, X, 1) = Chr(Asc(Mid(Secret, X, 1)) Xor Char)
      Next
      Encrypt = Secret
      'MsgBox Secret
End Function

Public Function Decrypt(Secret As Variant, CryptKey As Variant) As String
    'decrypt words from data file
    'secret = the string you wish to encrypt or decrypt.
    'CryptKey = the password (EKey=94022) used to encrypt the string.
    Dim L%, X%, Char%

    L = Len(CryptKey)
    For X = 1 To Len(Secret)
        Char = Asc(Mid(CryptKey, (X Mod L) - L * ((X Mod L) = 0), 1))
        Mid(Secret, X, 1) = Chr(Asc(Mid(Secret, X, 1)) Xor Char)
    Next
    Decrypt = Secret
    'MsgBox Secret
End Function
