VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDsPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oDSrc As New CDataSrc

'constructor

Private Sub Class_Initialize()

    oDSrc.Initialize 2, _
            sUId:=g_parm.getRParm(gc_sUID), _
            sUPwd:=g_parm.getRParm(gc_sPWD), _
            sUServer:=g_parm.getRParm(gc_sSERVER), _
            sUDB:=g_parm.getRParm(gc_sPADB), _
            sUDriver:="{SQL Server}"
    
End Sub

'destructor

Private Sub Class_Terminate()

    Set oDSrc = Nothing

End Sub

'accessor

Public Function getEnv() As rdoEnvironment
    
    Set getEnv = oDSrc.getEnv

End Function

'accessor

Public Function getCon() As rdoConnection

    Set getCon = oDSrc.getCon
    
End Function

Public Function OpenDs() As Boolean

    OpenDs = oDSrc.OpenDs
    
End Function

Public Sub CloseDs()

    oDSrc.CloseDs
    
End Sub

Public Function isOpened() As Boolean

    isOpened = oDSrc.isOpened

End Function


'execute SQL with no resultset expected

Public Function ExecSql(sSQL As String, pvOpt As Variant) As Boolean

    ExecSql = oDSrc.ExecSql(sSQL, pvOpt)

End Function

'begin transaction

Public Sub beginTrx()

    oDSrc.beginTrx

End Sub

'commit transaction

Public Sub commitTrx()

    oDSrc.commitTrx
    
End Sub

'rollback transaction

Public Sub rollbackTrx()

    oDSrc.rollbackTrx
    
End Sub

'get error message

Public Function getError() As String
    getError = oDSrc.getError
End Function

