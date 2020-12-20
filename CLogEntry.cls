VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'class implements log entry


Public dtDate As Date
Public iType As Integer
Public sMsg As String

Public Sub Initialize(pdtDate As Date, piType As Integer, psMsg As String)

    dtDate = pdtDate
    iType = piType
    sMsg = psMsg

End Sub


