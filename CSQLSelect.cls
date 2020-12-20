VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSQLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'module implements SQL Select data object


'pointer to a data source
Private oDS As Object

'sql string to be executed
Private sSql As String

'result set
Private oRs As rdoResultset

'error message from last execute
Private sError As String

'is recordset empty
Private bIsEmpty As Boolean

'constructor

Private Sub Class_Initialize()

    sSql = ""
    Set oRs = Nothing
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

    If Not oRs Is Nothing Then
        oRs.Close
    End If
    
    Set oRs = Nothing
    
End Sub

'run query

Public Function execute(Optional pvType As Variant = rdOpenForwardOnly, _
    Optional pvLockType As Variant = rdConcurReadOnly, _
    Optional pvOption As Variant = rdExecDirect) As Boolean

    CloseDo
    
    On Error GoTo ERROR_execute
    
    If Not oDS.isOpened() Then
        If oDS.OpenDs = False Then
            sError = oDS.getError()
            execute = False
            
            Exit Function
        End If
    End If
    
    Set oRs = oDS.getCon().OpenResultset(sSql, pvType, pvLockType, pvOption)

    If getEOF Then
        bIsEmpty = True
    Else
        bIsEmpty = False
    End If

    execute = True

    On Error Resume Next


    Exit Function

ERROR_execute:

    sError = Error(Err.Number)
    
    On Error Resume Next

    execute = False

End Function

Public Function getError() As String

    getError = sError

End Function

'return number of results
'not always available (-1)

Public Function getNumOfRows() As Long

    If oRs Is Nothing Then
        getNumOfRows = -1
    Else
        getNumOfRows = oRs.RowCount
    End If

End Function

'return status of query

Public Function isEmpty() As Boolean

    isEmpty = bIsEmpty

End Function


'access value of the column

Public Function getRsValue(iIndex As Integer) As Variant

    If oRs Is Nothing Or isEmpty Then
        getRsValue = Null
    Else
        getRsValue = oRs.rdoColumns.Item(iIndex).Value
    End If

End Function

'set new value for the collumn

Public Sub setRsValue(iIndex As Integer, vVal As Variant)

    If Not oRs Is Nothing Then
        oRs.rdoColumns.Item(iIndex).Value = vVal
    End If
    
End Sub

'access name of the collumn

Public Function getRsName(iIndex As Integer) As Variant

    If oRs Is Nothing Or isEmpty Then
        getRsName = Null
    Else
        getRsName = oRs.rdoColumns.Item(iIndex).Name
    End If
    
End Function

'update changes to resultset

Public Sub UpdateDo()

    If Not oRs Is Nothing And Not isEmpty Then
        oRs.Update
    End If
    
End Sub

'move to first record

Public Sub moveFirst()

    If Not oRs Is Nothing And Not isEmpty Then
        oRs.moveFirst
    End If
    
End Sub

'move to last record

Public Sub moveLast()

    If Not oRs Is Nothing And Not isEmpty Then
        oRs.moveLast
    End If
    
End Sub

'move to Prev record

Public Sub movePrev()

    If Not oRs Is Nothing And Not isEmpty Then
        oRs.MovePrevious
    End If

End Sub

'move to next record

Public Sub moveNext()

    If Not oRs Is Nothing And Not isEmpty Then
        oRs.moveNext
    End If
    
End Sub

'accesor

Public Function getEOF() As Boolean

    If Not oRs Is Nothing Then
        getEOF = oRs.EOF
    Else
        getEOF = True
    End If

End Function
