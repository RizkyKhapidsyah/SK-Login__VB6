VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oEnv As rdoEnvironment
Private oCon As rdoConnection


Private iMode As Integer

'id/pwd
Private sId As String
Private sPwd As String

'ODBC
Private sDSN As String

'direct
Private sServer As String
Private sDatabase As String
Private sDriver As String

Private sConStr As String

'errors
Private sError As String

'constructor

Private Sub Class_Initialize()

    Set oEnv = Nothing
    Set oCon = Nothing

End Sub

'destructor

Private Sub Class_Terminate()

    CloseDs
    Set oEnv = Nothing
    
End Sub


'first thing to call after creating new instance

Public Sub Initialize(ByVal iUMode As Integer, _
            Optional sUConStr As String = "", _
            Optional sUId As String = "", _
            Optional sUPwd As String = "", _
            Optional sUDsn As String = "", _
            Optional sUServer As String = "", _
            Optional sUDB As String = "", _
            Optional sUDriver As String = "")

'modes: 0 - connect string specified,
'       1 - odbc info provided,
'       2 - native transport info provided

    'save parameters
    iMode = iUMode
    sId = sUId
    sPwd = sUPwd
    sDSN = sUDsn
    sServer = sUServer
    sDatabase = sUDB
    sDriver = sUDriver
    sConStr = sUConStr

    'obtain default environment
    Set oEnv = rdoEngine.rdoEnvironments(0)
    
    'setup environment
    oEnv.CursorDriver = rdUseOdbc
    oEnv.LoginTimeout = 5

    'prepare connect string if not defined
    If iMode <> 0 Then
        sConStr = "UID=" & sId & ";PWD=" & sPwd & ";"
    End If

    'set up connection parameters
    Select Case iMode
    Case 0  'connect string
        sDSN = ""
    Case 1  'odbc
    Case 2  'direct
        sDSN = ""
        sConStr = sConStr & "driver=" & sDriver & ";Server=" & sServer _
            & ";database=" & sDatabase & ";"
    End Select
    
End Sub


Public Function OpenDs() As Boolean

    Set oCon = Nothing

    On Error GoTo ERROR_OpenDs

    'connect
    Set oCon = oEnv.OpenConnection(dsname:=sDSN, _
            Prompt:=rdDriverNoPrompt, _
            Connect:=sConStr)
            'Options:=rdAsyncEnable)

    'While oCon.StillConnecting
    '    DoEvents
    'Wend

    OpenDs = True

    Exit Function
    
ERROR_OpenDs:

    sError = Error(Err.Number)
    
    On Error Resume Next

    OpenDs = False
    
    Set oCon = Nothing

End Function

Public Sub CloseDs()
    
    If Not oCon Is Nothing Then
        oCon.Close
        Set oCon = Nothing
    End If
    
    Set oCon = Nothing
    
End Sub

' is connected ??

Public Function isOpened() As Boolean

    If oCon Is Nothing Then
        isOpened = False
    Else
        isOpened = True
    End If
    
End Function

'accessor to environment

Public Function getEnv() As rdoEnvironment

    Set getEnv = oEnv

End Function

'accessor to connection

Public Function getCon() As rdoConnection

    Set getCon = oCon
    
End Function

'execute SQL with no resultset expected

Public Function ExecSql(sSQL As String, pvOptions As Variant) As Boolean

    On Error GoTo ERROR_ExecSql

    If Not oCon Is Nothing Then
        oCon.execute sSQL, pvOptions
    End If

    ExecSql = True

    Exit Function

ERROR_ExecSql:

    sError = Error(Err.Number)
    
    On Error Resume Next

    ExecSql = False
    
End Function

'begin transaction

Public Sub beginTrx()

    oCon.BeginTrans

End Sub

'commit transaction

Public Sub commitTrx()

    oCon.CommitTrans
        
End Sub

'rollback transaction

Public Sub rollbackTrx()

    oCon.RollbackTrans
    
End Sub

'return error test

Public Function getError() As String
    getError = sError
End Function
