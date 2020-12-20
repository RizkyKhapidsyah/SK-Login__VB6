VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' tells if application is setup

Public Function isSetUp() As Boolean

    If g_parm.getRParm(gc_sCOMPLETE) = "" Then
    'If g_parm.getRParm(gc_sCOMPLETE) = "" Then
        isSetUp = False
    Else
        isSetUp = True
    End If
    
End Function

' reconnect with new settings
Public Sub reconnect()

    Set g_dsPA = Nothing

    Set g_dsPA = New CDsPA

End Sub

'converts string to Date
'accepted formats: yyyy, mm, dd, hh, nn, ss
'eg. toDate("date is: 1999.03.31 (13.45.59)", "date is: yyyy.mm.dd (hh.nn.ss)")

Public Function toDate(sDate As String, sFormat As String) As Date

    Dim sNewDate As String
    Dim sNewTime As String
    
    Dim dtDate As Date
    Dim tmTime As Date
    
    
    Dim iCYear As Integer
    Dim iCMonth As Integer
    Dim iCDay As Integer

    Dim iYearPos As Integer
    Dim iMonthPos As Integer
    Dim iDayPos As Integer

    Dim iCHour As Integer
    Dim iCMinute As Integer
    Dim iCSecond As Integer

    Dim iHourPos As Integer
    Dim iMinutePos As Integer
    Dim iSecondPos As Integer

    
    dtDate = Date
    tmTime = Time
    
    sFormat = LCase(sFormat)

    
    iCYear = Val(Format(dtDate, "yyyy"))
    iCMonth = Val(Format(dtDate, "mm"))
    iCDay = Val(Format(dtDate, "dd"))
    
    iCHour = Val(Format(tmTime, "hh"))
    iCMinute = Val(Format(tmTime, "nn"))
    iCSecond = Val(Format(tmTime, "ss"))

    
    iYearPos = getNonNull(InStr(sFormat, "yyyy"), 0)
    iMonthPos = getNonNull(InStr(sFormat, "mm"), 0)
    iDayPos = getNonNull(InStr(sFormat, "dd"), 0)
    
    iHourPos = getNonNull(InStr(sFormat, "hh"), 0)
    iMinutePos = getNonNull(InStr(sFormat, "nn"), 0)
    iSecondPos = getNonNull(InStr(sFormat, "ss"), 0)

    
    If iMonthPos > 0 Then
        sNewDate = Mid(sDate, iMonthPos, 2) & "/"
    Else
        sNewDate = Trim(CStr(iCMonth)) & "/"
    End If
    
    If iDayPos > 0 Then
        sNewDate = sNewDate & Mid(sDate, iDayPos, 2) & "/"
    Else
        sNewDate = sNewDate & Trim(CStr(iCDay)) & "/"
    End If
    
    If iYearPos > 0 Then
        sNewDate = sNewDate & Mid(sDate, iYearPos, 4)
    Else
        sNewDate = sNewDate & Trim(CStr(iCYear))
    End If


    If iHourPos > 0 Then
        sNewTime = Mid(sDate, iHourPos, 2) & ":"
    Else
        sNewTime = Trim(CStr(iCHour)) & ":"
    End If
    
    If iMinutePos > 0 Then
        sNewTime = sNewTime & Mid(sDate, iMinutePos, 2) & ":"
    Else
        sNewTime = sNewTime & Trim(CStr(iCMinute)) & ":"
    End If
    
    If iSecondPos > 0 Then
        sNewTime = sNewTime & Mid(sDate, iSecondPos, 2)
    Else
        sNewTime = sNewTime & Trim(CStr(iCSecond))
    End If

    
    toDate = sNewDate & " " & sNewTime
    
End Function


' returns non null value from the variant

Public Function getNonNull(pvVar As Variant, pvDefault As Variant)

    If IsNull(pvVar) Then
        getNonNull = pvDefault
    Else
        getNonNull = pvVar
    End If

End Function

'retyrn application path
Public Function getExePath() As String

    getExePath = App.Path
    
End Function

'replace single quote with MSSQL sentence
Public Function quote2MSSQL(psArg As String) As String

    Dim n As Integer
    Dim sRet As String
    Dim bDoQuote As Boolean
    Dim bFirstSwitch As Boolean
    
    
    
    If InStr(psArg, "'") <= 0 Then
        quote2MSSQL = " '" & psArg & "' "
        Exit Function
    End If
    
    
    
    sRet = " "
    bDoQuote = True
    bFirstSwitch = True
    
    For n = 1 To Len(psArg)
    
        If Mid(psArg, n, 1) = "'" Then
        
            
            If Not bDoQuote Then sRet = sRet & "'"
            
            If bFirstSwitch Then
                bFirstSwitch = False
            Else
                sRet = sRet & " + "
            End If
            
            sRet = sRet & " CHAR(39) "
            
            bDoQuote = True
        
        Else
        
       
            If bDoQuote Then
            
                If bFirstSwitch Then
                    bFirstSwitch = False
                Else
                    sRet = sRet & " + "
                End If
            
                sRet = sRet & "'"
            End If
        
            bDoQuote = False
           
            sRet = sRet & Mid(psArg, n, 1)
            
        End If
    
    
    
    Next n

    If Not bDoQuote Then
        sRet = sRet & "' "
    End If


    quote2MSSQL = sRet

End Function



