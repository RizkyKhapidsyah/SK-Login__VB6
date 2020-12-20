VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bInfoAdded As Boolean
Private bWarnAdded As Boolean
Private bErrorAdded As Boolean

Private sError As String
Private bIsError As Boolean


Private Sub Class_Initialize()
    resetStat
End Sub

'get error

Public Function isError(ByRef psError) As Boolean

    psError = sError
    isError = bIsError

End Function

Public Sub resetStat()

    bInfoAdded = False
    bWarnAdded = False
    bErrorAdded = False

End Sub

' test status functions

Public Function anyInfo() As Boolean
    anyInfo = bInfoAdded
End Function
Public Function anyWarning() As Boolean
    anyWarning = bWarnAdded
End Function
Public Function anyError() As Boolean
    anyError = bErrorAdded
End Function

'adds new entry in log file

Public Sub addEntry(piType As Integer, psMsg As String)

    Dim sFile As String
    Dim iFN As Integer

    bIsError = True
    iFN = -1
    
    On Error GoTo ERROR_add
    
    'retrieve or set if needed log file path
    
    If g_parm.getRParm(gc_sLOG, "") = "" Then
        
        g_parm.setRParm gc_sLOG, "c:\mysql.log"
        
        If g_parm.isError Then
            Exit Sub
        End If
    End If

    sFile = g_parm.getRParm(gc_sLOG)


    'generate new file number

    iFN = FreeFile
    
    'open log file
    
    Open sFile For Append Access Write Lock Write As #iFN

    Write #iFN, Format(Date, "yyyymmdd"), Format(Time, "hhnnss"), piType, psMsg

    Close #iFN


    Select Case piType
    Case gc_iLogInfo
        bInfoAdded = True
    Case gc_iLogWarning
        bWarnAdded = True
    Case gc_iLogError
        bErrorAdded = True
    End Select


    Exit Sub
    
ERROR_add:
    
    sError = Error(Err.Number)
    bIsError = True
    
    On Error Resume Next

    If iFN > -1 Then
        Close #iFN
    End If

End Sub

'returns log entries starting from the given date or all if null
' format is Date

Public Function getEntries(pvStart As Variant) As Collection

    Dim colRet As New Collection
    Dim oLogEntry As CLogEntry
    
    Dim iFN As Integer
    Dim sFile As String
    
    Dim sStart As String ' start date
    Dim dtStart As Date
    
    Dim sDate As String
    Dim sTime As String
    Dim iType As Integer
    Dim sMsg As String

    bIsError = False
    iFN = -1
    
    
    If Not IsNull(pvStart) Then
        dtStart = pvStart
        sStart = Format(dtStart, "yyyymmddhhnnss")
    End If
    
    
    On Error GoTo ERROR_getEntries
    
    'if file is not setup return empty collection
    
    If g_parm.getRParm(gc_sLOG, "") = "" Then
        
        Set getEntries = colRet
        
        Exit Function
    End If
        
    sFile = g_parm.getRParm(gc_sLOG)

    'open file
    
    iFN = FreeFile
    Open sFile For Input Access Read As #iFN


    While Not EOF(iFN)


        Input #iFN, sDate, sTime, iType, sMsg



        'if no start date - return all log filr
        
        If IsNull(pvStart) Then
        
            Set oLogEntry = New CLogEntry
            oLogEntry.Initialize g_util.toDate(sDate & sTime, "yyyymmddhhnnss"), _
                    iType, sMsg
        
            colRet.Add oLogEntry
        
        Else
            'return only requested entries
            
            If sDate & sTime >= sStart Then
            
                Set oLogEntry = New CLogEntry
                oLogEntry.Initialize g_util.toDate(sDate & sTime, "yyyymmddhhnnss"), _
                    iType, sMsg
        
                colRet.Add oLogEntry
            
            End If
        End If
        
    Wend
    
    Close #iFN
    
    Set getEntries = colRet
    
    Exit Function
    
ERROR_getEntries:
    
    sError = Error(Err.Number)
    bIsError = True
    
    On Error Resume Next

    If iFN > -1 Then
        Close #iFN
    End If

    Set getEntries = colRet
    
End Function


Public Sub clearContents()

  
    Dim iFN As Integer
    Dim sFile As String
    
    iFN = -1
    bIsError = False
    
    On Error GoTo ERROR_clearContents
    
    'if file is not setup return empty collection
    
    If IsNull(g_parm.getRParm(gc_sLOG)) Then
        
        Exit Sub
    End If
        
    sFile = g_parm.getRParm(gc_sLOG)

    'open file
    
    iFN = FreeFile
    Open sFile For Output Access Write Lock Write As #iFN
    Close #iFN

    Exit Sub
    
ERROR_clearContents:

    sError = Error(Err.Number)
    bIsError = True
    
    On Error Resume Next

    If iFN > -1 Then
        Close #iFN
    End If

End Sub
