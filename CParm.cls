VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   module is used to provide access to
'   Registry settings of HKEY_CURRENT_USER
'   Software, VB and VBA Program Settings, PAdb, Setup

Private sError As String
Private bError As Boolean

Public Function isError() As Boolean
    isError = bError
End Function

Public Function getError() As String
    getError = sError
End Function

'-------------------------------------------------------
'REGISTRY lookups
'-------------------------------------------------------

'get registry parameter

Public Function getRParm(psKey As String, Optional psDefault As String = "") As String

    getRParm = _
        GetSetting(appname:="PAdb", section:="Setup", _
            Key:=psKey, Default:=psDefault)

End Function

'set registry parameter

Public Sub setRParm(psKey As String, psVal As String)

    SaveSetting appname:="PAdb", section:="Setup", _
        Key:=psKey, setting:=psVal

End Sub

'remove all registry parameters

Public Sub removeRAllParm()

    DeleteSetting "PAdb", "Setup"

End Sub

'remove registry parameter

Public Sub removeRParm(psKey As String)

    DeleteSetting "PAdb", "Setup", psKey

End Sub
