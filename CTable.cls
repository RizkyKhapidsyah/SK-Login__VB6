VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sError As String
Private bError As Boolean
Private dtRunDate As Date
Private dtPayPeriod As Date

Public Function isError() As Boolean
    isError = bError
End Function

Public Function getError() As String
    getError = sError
End Function
Public Function Create() As Boolean
    'create dbuser table

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

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Drop dbuser table"
    sSQL = "if exists " & _
               "(select * from sysobjects where " & _
               "id = object_id(N'[dbo].[dbuser]') and " & _
               "OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
               "drop table [dbo].[dbuser]"

    dosPaDel.setSql sSQL

    If dosPaDel.execute = False Then
        g_log.addEntry gc_iLogError, dosPaDel.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If
    
    'MsgBox "2Create DbUuser table"
    sSQL = "create table dbo.dbuser([uid] [varchar] (12) NULL ," & _
               "[status] [varchar] (2) NULL ," & _
               "[name] [varchar] (50) NULL)"

    dospains.setSql sSQL

    If dospains.execute = False Then
        g_log.addEntry gc_iLogError, dospains.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If
    
    'Msgbox "3Check for user table"
    sSQL = "select name from sysobjects"
    dospasel.setSql sSQL

    If dospasel.execute = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If

    bFound = False
    While Not dospasel.getEOF And bFound = False
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
        If tmpstr = "dbuser" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
        End If
        dospasel.moveNext
    Wend
    If Not bFound Then
        gDbuser = True
        g_log.resetStat
        g_log.addEntry gc_iLogInfo, "dbuser table was created."
    End If
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "dbuser table was created."
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
Public Function Edit() As Boolean
    'edit table
End Function

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

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Drop tblusers table"
'    sSQL = "if exists " & _
'               "(select * from sysobjects where " & _
'               "id = object_id(N'[dbo].[tblUsers]') and " & _
'               "OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
'               "drop table [dbo].[tblUsers]"
'
'    dosPaDel.setSql sSQL
'
'    If dosPaDel.execute = False Then
'        g_log.addEntry gc_iLogError, dosPaDel.getError
'        bRet = False
'        'MsgBox "Error here1"
'        GoTo ERROR_run
'    End If
    
    'MsgBox "2Create tblusers table"
    sSQL = "create table dbo.tblusers([UserID] [varchar] (12) NULL ," & _
               "[Password] [varchar] (15) NULL ," & _
               "[UserGroup] [varchar] (20) NULL ," & _
               "[ActivationDate] [datetime] NULL ," & _
               "[ExpirationDate] [datetime] NULL ," & _
               "[UserName] [varchar] (20) NULL)"

    dospains.setSql sSQL

    If dospains.execute = False Then
        g_log.addEntry gc_iLogError, dospains.getError
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
    While Not dospasel.getEOF And bFound = False
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
        If tmpstr = "tblusers" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
        End If
        dospasel.moveNext
    Wend
    gFound = False
    If bFound = True Then
        'tblusers table was found, need to insert admin record
        sSQL = "insert into tblusers (userid, password, usergroup, activationdate, expirationdate, username)" _
             + "values ('admin', ' ', 'Administrator', '05/01/2001', '05/01/2002', 'Administrator')"

        dospains.setSql sSQL

        If dospains.execute = False Then
            g_log.addEntry gc_iLogError, dospains.getError
            bRet = False
            'MsgBox "Error here1"
            GoTo ERROR_run
        End If
        g_log.resetStat
        g_log.addEntry gc_iLogInfo, "User table was created."
        gFound = True
    Else
        Msg = "Error, please call Administrator!"
        MsgBox Msg, 16, "Login Error"
    End If
   
    'use this to execute a stored procedure
'    MySqldoc = "LoadJbTable"
'    myexe = g_dsPA.ExecSql(MySqldoc, rdExecDirect)

'    If myexe = False Then
'        g_log.addEntry gc_iLogError, g_dsPA.getError
'        bRet = False
'        'MsgBox "Error here2"
'        GoTo ERROR_run
'    End If
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "User table was created."
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
Public Function Insert() As Boolean
    'insert into dbuser table
    
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

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Insert fields into dbuser table"
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
        If tmpstr = "dbuser" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
        End If
        dospasel.moveNext
    Wend
    If bFound = True Then
        'tblusers table was found, need to insert admin record
        sSQL = "insert into dbuser (uid, status, name)" _
             + "select uid, status, name from sysusers"

        dospains.setSql sSQL

        If dospains.execute = False Then
            g_log.addEntry gc_iLogError, dospains.getError
            bRet = False
            'MsgBox "Error here1"
            GoTo ERROR_run
        End If
        g_log.resetStat
        g_log.addEntry gc_iLogInfo, "Data was inserted into dbuser table."
    Else
        Msg = "Error, please call Administrator!"
        MsgBox Msg, 16, "Login Error"
        gDbuser = False
    End If
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "Data was inserted into dbuser table."
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
Public Function Drop() As Boolean
    'Drop dbuser table
    
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

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Drop dbUser table"
    sSQL = "if exists " & _
               "(select * from sysobjects where " & _
               "id = object_id(N'[dbo].[dbUser]') and " & _
               "OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
               "drop table [dbo].[dbUser]"

    dosPaDel.setSql sSQL

    If dosPaDel.execute = False Then
        g_log.addEntry gc_iLogError, dosPaDel.getError
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

    bFound = False
    While Not dospasel.getEOF And bFound = False
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
        If tmpstr = "dbuser" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
        End If
        dospasel.moveNext
    Wend
    If Not bFound Then
        gDbuser = True
        g_log.resetStat
        g_log.addEntry gc_iLogInfo, "dbUser table was found."
    End If
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "dbuser table was found."
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


Public Sub Delete()

    Dim dodParms As New CSQLDelete
        
    dodParms.Initialize g_dsPA
        
    dodParms.setSql " DELETE FROM DBUSER "
    
    If dodParms.execute = False Then
        bError = True
        sError = dodParms.getError
    Else
        bError = False
    End If
        
    Set dodParms = Nothing

End Sub

Public Function TableExist() As Boolean
    'search for table
    Dim dtLastRun As Date
    Dim tmpstr As String

    'return code
    Dim bRet As Boolean
  
    Dim sSQL As String  'sql string
    Dim bFound As Boolean
    
    Dim dsPA As New CDsPA       'ds for selects only
    
    'data objects
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    Dim dosPaDel As New CSQLDelete

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Insert fields into dbuser table"
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
    TableExist = False
    While Not dospasel.getEOF And bFound = False
        tmpstr = LCase(Trim(dospasel.getRsValue(0)))
        If tmpstr = "dbuser" Then
            'MsgBox dospasel.getRsValue(0) & ","
            bFound = True
            TableExist = True
        End If
        dospasel.moveNext
    Wend
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "Data was inserted into DbUser table."
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

Public Function IsEmptyTable() As Boolean
    'is the dbuser table empty
    Dim dtLastRun As Date
    Dim tmpstr As String

    'return code
    Dim bRet As Boolean
  
    Dim sSQL As String  'sql string
    Dim bFound As Boolean
    
    Dim dsPA As New CDsPA       'ds for selects only
    
    'data objects
    Dim dospasel As New CSQLSelect
    Dim dospains As New CSQLInsert
    Dim dosPaDel As New CSQLDelete

    dospasel.Initialize dsPA
    dospains.Initialize g_dsPA
    dosPaDel.Initialize g_dsPA
    
    ' initialization
    dtRunDate = Date + Time
    
    bRet = True
    
    'start your functionality here
    
    'MsgBox "1 Insert fields into dbuser table"
    sSQL = "select uid, name, status from dbuser"
    dospasel.setSql sSQL

    If dospasel.execute = False Then
        g_log.addEntry gc_iLogError, dospasel.getError
        bRet = False
        'MsgBox "Error here1"
        GoTo ERROR_run
    End If

    If dospasel.getEOF Then
        IsEmptyTable = True
    Else
        IsEmptyTable = False
    End If
    Exit Function
    
ERROR_run:
    If bRet = True Then
        g_log.addEntry gc_iLogInfo, "Data found in dbuser table."
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
