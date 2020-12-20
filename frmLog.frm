VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Log File"
   ClientHeight    =   4920
   ClientLeft      =   3432
   ClientTop       =   2964
   ClientWidth     =   6132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6132
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstLog 
      Height          =   4212
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5892
      _ExtentX        =   10393
      _ExtentY        =   7430
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Level"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Message"
         Object.Width           =   52917
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6132
      _ExtentX        =   10816
      _ExtentY        =   593
      ButtonWidth     =   487
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4920
      Top             =   5400
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vLogStart As Variant 'log start date as Date

Public oLog As Object 'log to use

Private Sub populateList()

    Dim sError As String
    Dim colLog As Collection
    Dim oLogEntry As CLogEntry
    Dim oLi As ListItem

    If oLog Is Nothing Then
        Exit Sub
    End If


    Set colLog = oLog.getEntries(vLogStart)

    If oLog.isError(sError) = True Then
        MsgBox sError, vbOKOnly, "APP Error"
        
        Exit Sub
    End If

    For Each oLogEntry In colLog

        Set oLi = lstLog.ListItems.Add(, , _
            CStr(Format(oLogEntry.dtDate, "mm/dd/yyyy hh:nn:ss")))
        
        'set text for second column
        Select Case oLogEntry.iType
        Case gc_iLogInfo
            oLi.SubItems(1) = "Info"
        Case gc_iLogWarning
            oLi.SubItems(1) = "Warning"
        Case gc_iLogError
            oLi.SubItems(1) = "Error"
        Case Else
            oLi.SubItems(1) = "Unknown"
        End Select

        oLi.SubItems(2) = oLogEntry.sMsg
        
    Next oLogEntry

End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    populateList

End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kCancel"
            HandleCancelClick
    End Select

End Sub

Private Sub HandleCancelClick()
    Unload frmLog
End Sub

