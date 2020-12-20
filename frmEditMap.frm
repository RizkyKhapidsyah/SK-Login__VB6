VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DETAILS"
   ClientHeight    =   1512
   ClientLeft      =   2736
   ClientTop       =   5136
   ClientWidth     =   5292
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1512
   ScaleWidth      =   5292
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tbrMain2 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5292
      _ExtentX        =   9335
      _ExtentY        =   593
      ButtonWidth     =   487
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   3852
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   3852
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   480
      Top             =   2280
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
            Picture         =   "frmEditMap.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditMap.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "UID:"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   720
   End
End
Attribute VB_Name = "frmEditMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oItem As ListItem   ' list item of the view

Private iMode As Integer    'editing mode
                            '0-edit
                            '1-insert
                            '2-delete
                            
Private doSql As Object     'sql object to be executed after
                            ' ok is selected
                            
                            
Public sName   As String    'used to return value back to parent
Public sId  As String       'used to return value back to parent


Private bResult As Boolean  ' result of record editing
                            '(true - action commited)/(false - disregard action)

Public Sub setEditMode(pMode As Integer, pItem As ListItem, _
        pSql As Object)

    Set oItem = pItem
    iMode = pMode
    Set doSql = pSql
    
End Sub

'query return status

Public Function getResult() As Boolean
    
    getResult = bResult
    
End Function

Private Sub cmdOK_Click()

    Dim vID As String
    
    vID = txtId.Text
    
    If IsNull(vID) Then
        Exit Sub
    End If

    sName = txtName.Text    'UCase(txtName.Text)
    sId = vID

    If sName <> "" Then

        Select Case iMode
        
        Case 0  'edit
        
            doSql.setSql "UPDATE Dbuser " _
                & " SET name = '" & sName & "' " _
                & " WHERE uid = '" & sId & "' "

        Case 1  'insert
        
            'doSql.setSql "INSERT INTO jbcustomer (custid, custname) " _
            '    & " VALUES ('" & sID & "', " _
            '    & " '" & sName & "') "
        
        Case 2  'delete

            doSql.setSql "DELETE FROM Dbuser " _
                & " WHERE uid = '" & sId & "' "

        End Select

        bResult = doSql.execute
        
        If bResult = False Then
            MsgBox doSql.getError, vbOKOnly, "SQL Error"
        End If
        
    End If
    
    If bResult = True Then
        Unload Me
    End If
    
End Sub


Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    bResult = False

    Select Case iMode
    
    Case 0  'edit
    
        Me.Caption = "Edit Mapping"
        
        txtName.Text = oItem.SubItems(1)
        txtId = oItem.Text
    
        txtName.Enabled = True
        txtId.Enabled = False

    Case 1  'insert
    
'        Me.Caption = "Insert Mapping"

'        txtName.Enabled = True
'        txtId.Enabled = True

    Case 2  'delete
    
        Me.Caption = "Confirm Delete"
        
        txtName.Text = oItem.SubItems(1)
        txtId = oItem.Text

        txtName.Enabled = False
        txtId.Enabled = False

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set doSql = Nothing

End Sub

Private Sub tbrMain2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kFinish2"
            cmdOK_Click
        Case "kCancel2"
            'clear button
            Unload Me
        Case "kHelp2"
            MsgBox "Display Help Screen.", vbInformation
    End Select

End Sub


