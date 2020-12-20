Attribute VB_Name = "MMain"
Option Explicit
'project-wide global variable

'data sourcess
Public g_dsPA As New CDsPA      'data src for pa

    'log object
Public g_log As New CLog

    'parameter access objects
Public g_parm As New CParm      'interface for parameters

    'utility cladd
Public g_util As New CUtils

    'task objects
'declare class modules here
Public g_oWip As New CWip       'transfer data from one database to another
Public g_oTable As New CTable   'create user table

    'status variables
Public g_iRunMode As Integer    ' run mode:
                                ' 0 - window-less
                                ' 1 non window less


'constants
Public Const gc_iString As Integer = 0
Public Const gc_iNumber As Integer = 1
Public Const gc_iDate As Integer = 2

'setup screen parameters
Public Const gc_sSERVER As String = "SERVER"
Public Const gc_sPADB As String = "PADB"
Public Const gc_sUID As String = "UID"
Public Const gc_sPWD As String = "PWD"

Public Const gc_sLOG As String = "LOGFILE"
Public Const gc_sCOMPLETE As String = "COMPLETE"

    'log related
Public Const gc_iLogInfo As Integer = 0
Public Const gc_iLogWarning As Integer = 1
Public Const gc_iLogError As Integer = 2

Global DgDef As Integer
Global Response As Integer
Global Style As String
Global Title As String
Global Help As String
Global Msg As String
Global Ans As String
Global Default As String
Global gStartDate As Variant
Global gEndDate As Variant
Global gWkendDate As Variant
Global badlogin As Integer
Global gDbuser As Integer

'messageBox
Global Const HM_MSG = "MST"    ' used in title of msgbox
Global Const MB_OK = 0, MB_OKCANCEL = 1     ' define buttons
Global Const MB_YESNOCANCEL = 3, MB_YESNO = 4
Global Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32    ' define icons
Global Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
Global Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7, IDCANCEL = 2, IDOK = 1   ' define other
Global Const MB_DEFBUTTON1 = 0, MB_DEFBUTTON3 = 512
' Message box beep constants
Global Const MB_DEFBEEP = -1
Global Const MB_ICONASTERISK = 64
Global Const MB_ICONHAND = 16
Public Function Validate_WkDate(WkendDate)
    ' procedure to validate week ending date

    Dim DayNum As Integer

    Validate_WkDate = False
    If IsDate(WkendDate) Then
        DayNum = Weekday(WkendDate)
        ' see if it's a weekend day or a weekday
        If DayNum = 1 Then
            Validate_WkDate = True
        End If
    End If
    'MsgBox "DayNum must be a ONE for Sunday: " & DayNum
    'MsgBox "DayNum must be a SIX for Friday: " & DayNum

End Function

Public Function Validate_StartDate(StartDate)
    ' procedure to validate week ending date

    Dim DayNum As Integer

    Validate_StartDate = False
    If IsDate(StartDate) Then
        DayNum = Weekday(StartDate)
        ' must be a Monday
        If DayNum = 2 Then
            Validate_StartDate = True
        End If
    End If
    'MsgBox "DayNum must be a TWO for Sunday: " & DayNum

End Function
' start point of entire application
Public Sub Main()
    
    Dim bGUI As Boolean
    
    g_iRunMode = 1  'window mode
                    'will do nothing but setup if needs
    bGUI = True
    
    ' check if app is setup
    If g_util.isSetUp = False And g_iRunMode <> 0 Then
        frmSetup.Show vbModal
    End If
    ' if it is still not setup - exit
    If g_util.isSetUp = False Then
        If g_iRunMode <> 0 Then
            MsgBox "Application is not setup. Cant run.", vbOKOnly, "APP Error"
        End If
        Exit Sub
    End If

    If bGUI Then
        badlogin = True
        gLogin = False
        frmLogin.Show vbModal
        If Not badlogin Then
            'frmConsole.Show vbModal
            frmControl.Show vbModal
        End If
    End If

End Sub
