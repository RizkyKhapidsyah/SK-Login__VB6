VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSQLDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'module implements SQL Delete data object


'pointer to a data source
Private oDS As Object

'sql string to be executed
Private sSql As String

'error message from last execute
Private sError As String



'constructor

Private Sub Class_Initialize()

    sSql = ""
    Set oDS = Nothing

End Sub

'destructor

Private Sub Class_Terminate()

    CloseDo
    Set oDS = Nothing
    
End Sub

'assign ds to sql object

Public Sub Initialize(oUDS As Object)

    Set oDS = oUDS

End Sub

'assign sql to execute

Public Sub setSql(sUSql As String)

    sSql = sUSql

End Sub

'close SQL

Public Sub CloseDo()

    
End Sub

'run query

Public Function execute() As Boolean

    CloseDo
    
    On Error GoTo ERROR_execute
    
    If Not oDS.isOpened() Then
        If oDS.OpenDs = False Then
            sError = oDS.getError()
            execute = False
            
            Exit Function
        End If
    End If

    execute = oDS.ExecSql(sSql, rdExecDirect)

    If execute = False Then
        sError = oDS.getError()
    End If


    Exit Function

ERROR_execute:

    sError = Error(Err.Number)
    
    On Error Resume Next

    execute = False

End Function

Public Function getError() As String

    getError = sError

End Function

