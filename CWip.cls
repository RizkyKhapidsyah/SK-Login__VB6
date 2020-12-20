VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sError As String
Private dtRunDate As Date
Private dtPayPeriod As Date


Public Function run() As Boolean

    'times of run
    Dim dtLastRun As Date
    Dim tmpstr As String
    Dim MySqldoc As String
    Dim myexe As Integer

    'return code
    Dim bRet As Boolean
  
    Dim sSQL As String  'sql string
    Dim bFound As Boolean
    
    Dim dsPA As New CDsPA       'ds for selects only
    
    'data objects
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    Dim dosPaDel As New CSQLDelete
    Dim dosAppSel As New CSQLSelect
    Dim dosAppIns As New CSQLInsert
    Dim dosAppDel As New CSQLDelete

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
'    dtPayPeriod = DateValue(frmWipRun.txtDate.Text)
    
    run = True
    bRet = True
    
    g_log.resetStat
    g_log.addEntry gc_iLogInfo, "Data Transfer Started"
    
    If g_iRunMode <> 0 Then
        frmConsole.MousePointer = vbHourglass
    End If
    
    'start your functionality here
     
    'MsgBox "1 Clean Customer Table"
'    sSQL = "delete from jbcustomer"
'    dosPaDel.setSql sSQL

'    If dosPaDel.execute = False Then
'        g_log.addEntry gc_iLogError, dosPaDel.getError
'        bRet = False
'        'MsgBox "Error here1"
'        GoTo ERROR_run
'    End If
    
    sSQL = "select userid, password from tblusers"
    dospasel.setSql sSQL

    If dospasel.execute = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If
    
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
    While Not dospasel.getEOF
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
'        MsgBox tmpstr & ","
        If tmpstr = "login" Then
            MsgBox dospasel.getRsValue(0) & ","
            bFound = True
        '    Exit Do
        End If
        dospasel.moveNext
    Wend
    If bFound = False Then
        'login table was not found, create login table
        MsgBox "Error - table not found"
    Else
        MsgBox "Table found!"
    End If
   

    'execute a stored procedure
    MySqldoc = "LoadJbTable"
'    myexe = g_dsPA.ExecSql(MySqldoc, rdExecDirect)

'    If myexe = False Then
'        g_log.addEntry gc_iLogError, g_dsPA.getError
'        bRet = False
'        'MsgBox "Error here2"
'        GoTo ERROR_run
'    End If

ERROR_run:
    
    run = bRet
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "Data Transfer - Process Complete!"
    End If
    
    ' show log
    If g_iRunMode <> 0 Then
    
        frmConsole.MousePointer = vbDefault

        Set frmLog.oLog = g_log
        frmLog.vLogStart = dtRunDate
        frmLog.Show vbModal
        
    End If

    g_log.resetStat

End Function
