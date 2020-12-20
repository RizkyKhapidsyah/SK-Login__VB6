VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWEndg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Week Ending Date"
   ClientHeight    =   1788
   ClientLeft      =   2724
   ClientTop       =   2760
   ClientWidth     =   4608
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1788
   ScaleWidth      =   4608
   Begin MSComctlLib.Toolbar tbrMain2 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4608
      _ExtentX        =   8128
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kFinish2"
            Object.ToolTipText     =   "Finish"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel2"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep2"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kHelp2"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Date"
      Height          =   372
      Left            =   2520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1092
   End
   Begin VB.TextBox txtEndDate 
      BackColor       =   &H00C0C0C0&
      Height          =   288
      Left            =   2880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   972
   End
   Begin VB.TextBox txtStartDate 
      BackColor       =   &H00C0C0C0&
      Height          =   288
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton cmdWkDate 
      Caption         =   "Click to Enter Date"
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   1452
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   2520
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
            Picture         =   "frmWEndg.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWEndg.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1092
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payroll Period - Monday through Sunday"
      Height          =   1092
      Left            =   720
      TabIndex        =   9
      Top             =   2040
      Width           =   3372
      Begin VB.Label lblthru 
         Alignment       =   2  'Center
         Caption         =   "through"
         Height          =   252
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   732
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Monday:"
      Height          =   252
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sunday:"
      Height          =   252
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   972
   End
   Begin VB.Label lbldate 
      Alignment       =   1  'Right Justify
      Caption         =   "Week Ending Date is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2052
   End
End
Attribute VB_Name = "frmWEndg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dtDate As Date
Dim sDatebuild As String
Dim iPosition As Integer
Dim bdate As Boolean
Dim WkendDate As Variant
Dim StartDate As Variant
Dim EndDate As Variant
Dim Msg1 As String, Msg2 As String

Public Sub getwkdate(flag As Integer)
    ' prompt user for weekending date
    ' called from WeekEndDate_Click

    Dim Tmpflag As Integer, TempFlag As Integer
    Dim Msg1 As String, Msg2 As String
    Dim OldWkEndDate As Variant
    Dim iPosition As Integer
    Dim dtDate As Date
    Dim sDatebuild As String
    Dim bdate As Boolean
    
    Tmpflag = True
    OldWkEndDate = WkendDate
    iPosition = 0
    bdate = False
    flag = False
    Do While Tmpflag
        Msg1 = "Please enter Week Ending Date"
        Msg1 = Msg1 & " in the form mm/dd/yyyy."
        WkendDate = InputBox$(Msg1, HM_MSG, WkendDate)
        If WkendDate <> "" Then
            iPosition = InStr(WkendDate, "/")
            If iPosition > 0 Then
                If IsDate(WkendDate) Then
                    WkendDate = CDate(WkendDate)
                    bdate = True
                    'MsgBox "position > 0"
                End If
            Else
                'If Not Len(WkendDate) < 6 Then
                If Len(WkendDate) = 6 Or Len(WkendDate) = 8 Then
                    sDatebuild = Left(WkendDate, 2) + "/" + Mid(WkendDate, 3, 2) + "/" + Right(WkendDate, Len(WkendDate) - 4)
                    If IsDate(sDatebuild) Then
                        bdate = True
                        WkendDate = CDate(sDatebuild)
                        'MsgBox "position < 6"
                    End If
                Else
                    bdate = False
                End If
            End If
            If bdate Then
                If DateValue(WkendDate) < DateValue("1/1/1998") Then bdate = False
                If DateValue(WkendDate) > DateValue("1/1/2050") Then bdate = False
            End If
            TempFlag = True
            If bdate And TempFlag = Validate_WkDate(WkendDate) Then
                Tmpflag = False
                'If Len(WkendDate) = 8 Then
                '    WkendDate = Format(WkendDate, "mm/dd/yyyy")
                'End If
                txtDate.Text = WkendDate
                Tmpflag = False
                flag = True
            Else
                DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2
                Msg2 = "Invalid Week Ending Date! " & WkendDate
                MsgBox Msg2, DgDef, HM_MSG
                Tmpflag = False     ' exit loop
                WkendDate = ""
                txtDate.Text = WkendDate
                Exit Do
            End If
            'input box is empty do nothing
        Else
            WkendDate = OldWkEndDate
            Tmpflag = False
        End If
    Loop
End Sub     ' GetWkDate
Private Sub GetStartDate(flag As Integer)
    'call to get Monday and Sunday dates
    Dim TempFlag As Integer

    StartDate = CDate(WkendDate) - 6
    EndDate = WkendDate
    TempFlag = True
    If TempFlag = Validate_StartDate(StartDate) Then
        flag = True
    Else
        DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2
        Msg2 = "Invalid Start Date! " & StartDate
        MsgBox Msg2, DgDef, HM_MSG
        flag = False
    End If
End Sub

Private Sub cmdClear_Click()
    'clear date fields
    txtStartDate.Text = ""
    txtEndDate.Text = ""
    txtDate.Text = ""
    WkendDate = ""
    StartDate = ""
    EndDate = ""
End Sub

Private Sub cmdWkDate_Click()
    'prompt user for w/e date
    Dim flag As Integer
    Dim bdate As Boolean
   
    'get date from textbox
    flag = False
    bdate = True
    'call calendar
    frmCalendar.Show vbModal
    WkendDate = gWkendDate
'    getwkdate flag
    If Validate_WkDate(WkendDate) Then
        txtDate.Text = WkendDate
        If DateValue(WkendDate) < DateValue("1/1/1998") Then bdate = False
        If DateValue(WkendDate) > DateValue("1/1/2050") Then bdate = False
        If bdate Then
            GetStartDate flag
        End If
    Else
        DgDef = MB_OK + MB_ICONEXCLAMATION + MB_DEFBUTTON2
        Msg2 = "Invalid Week Ending Date!  " & WkendDate
        MsgBox Msg2, DgDef, HM_MSG
        txtDate.Text = ""
        txtStartDate.Text = ""
        txtEndDate.Text = ""
        flag = False
    End If

    If flag Then
        'w/e date, StartDate, and EndDate are valid
        txtStartDate.Text = StartDate
        txtEndDate.Text = EndDate
        'cmdWkDate.Enabled = False
    End If
End Sub


Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    'txtDate.SetFocus
    ' Start highlight before first character.
    txtDate.SelStart = 0
    ' Highlight to end of text.
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub tbrMain2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kFinish2"
            HandleFinish2Click
        Case "kCancel2"
            'clear button
            Unload Me
        Case "kHelp2"
            MsgBox "Display Help Screen.", vbInformation
    End Select

End Sub
Private Sub HandleFinish2Click()
    'MsgBox "Finish.", vbInformation
'    Msg = "Do you want to continue?"
'    Style = vbYesNo + vbCritical + vbDefaultButton2
'    Ans = MsgBox(Msg, Style, HM_MSG)
    'ans will return as IDYES = 6 or IDNO = 7
'    If Ans = IDYES Then
        If Not IsNull(txtDate) And txtDate <> Empty Then
            gStartDate = CDate(StartDate)
            gEndDate = CDate(EndDate)
            gWkendDate = CDate(WkendDate)
        Else
            gStartDate = ""
            gEndDate = ""
            gWkendDate = ""
        End If
        'could execute a procedure to perform some other task
        Unload Me
'    End If  'answer is yes
End Sub
