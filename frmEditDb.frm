VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditDb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Table"
   ClientHeight    =   4320
   ClientLeft      =   3180
   ClientTop       =   2136
   ClientWidth     =   5268
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5268
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstParms 
      Height          =   3132
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   5525
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain2 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5268
      _ExtentX        =   9292
      _ExtentY        =   593
      ButtonWidth     =   487
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep1"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kCancel2"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kSep2"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "kHelp2"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3480
      Top             =   4680
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
            Picture         =   "frmEditDb.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditDb.frx":1DF4
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
End
Attribute VB_Name = "frmEditDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmEdit As frmEditMap

'populate list of parameters

Private Function populateList() As Boolean

    Dim dosParms As New CSQLSelect  ' data object to retreive params
    Dim oLi As ListItem             ' pointer to inserted item
    
    populateList = True
   
    dosParms.Initialize g_dsPA

    dosParms.setSql "SELECT uid, name, status FROM dbuser"

    If dosParms.execute = False Then
        MsgBox dosParms.getError, vbOKOnly, "SQL Error"
        populateList = False
        Exit Function
    End If
    
    
    While Not dosParms.getEOF

        Set oLi = lstParms.ListItems.Add(, , Trim(CStr(dosParms.getRsValue(0))))

      
        'set text for second column

        oLi.SubItems(1) = Trim(CStr(dosParms.getRsValue(1)))
        oLi.SubItems(2) = Trim(CStr(dosParms.getRsValue(2)))

        dosParms.moveNext

    Wend

    Set dosParms = Nothing

End Function


' do editing of the record

Private Sub doEdit()
    
    Dim oLi As ListItem
    
    Dim douParms As New CSQLUpdate
    
    
    douParms.Initialize g_dsPA


    If Not lstParms.SelectedItem Is Nothing Then

        Set oLi = lstParms.SelectedItem

        Set frmEdit = New frmEditMap
        'set form to update mode
        frmEdit.setEditMode 0, oLi, douParms
        
        frmEdit.Show vbModal, Me
        
        If frmEdit.getResult = True Then
        
            oLi.Text = frmEdit.sId
            oLi.SubItems(1) = frmEdit.sName
            
            lstParms.Refresh
        
        End If
        
        Set frmEdit = Nothing

    End If
     
    Set douParms = Nothing

End Sub

' do inserting of the record

Private Sub doInsert()
    
    Dim doiParms As New CSQLInsert
    
    Dim oLi As ListItem
    
    
    doiParms.Initialize g_dsPA

    Set frmEdit = New frmEditMap
    'set form to insert mode
    frmEdit.setEditMode 1, Nothing, doiParms
    
    frmEdit.Show vbModal, Me
    
    If frmEdit.getResult = True Then
        
        'lstParms.ListItems.a oLi.Index
        
        Set oLi = lstParms.ListItems.Add(, , frmEdit.sId)
        oLi.SubItems(1) = frmEdit.sName
        
        lstParms.Refresh
       
    End If
    
    Set frmEdit = Nothing
        
    Set doiParms = Nothing
        
End Sub

' do deleting of the record

Private Sub doDelete()
    
    Dim oLi As ListItem
    
    Dim dodParms As New CSQLDelete
    
    
    dodParms.Initialize g_dsPA


    If Not lstParms.SelectedItem Is Nothing Then

        Set oLi = lstParms.SelectedItem

        Set frmEdit = New frmEditMap
        'set form to update mode
        frmEdit.setEditMode 2, oLi, dodParms
        
        frmEdit.Show vbModal, Me
        
        If frmEdit.getResult = True Then
        
            lstParms.ListItems.Remove oLi.Index
            
            lstParms.Refresh
        
        End If
        
        Set frmEdit = Nothing
        
    End If
 
    Set dodParms = Nothing
        
End Sub

Private Sub cmdDelete_Click()
    doDelete
End Sub

Private Sub cmdInsert_Click()
'    doInsert
End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    populateList
End Sub


Private Sub lstParms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    lstParms.SortKey = ColumnHeader.Index - 1
    
    ' Set Sorted to True to sort the list.
    lstParms.Sorted = True

End Sub

Private Sub lstParms_DblClick()
    doEdit
End Sub

Private Sub cmdEdit_Click()
    doEdit
End Sub

Private Sub tbrMain2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kCancel2"
            'clear button
            HandleCancel2Click
        Case "kHelp2"
            HandleHelp2Click
    End Select

End Sub
Private Sub HandleCancel2Click()
    Unload frmEditDb
End Sub

Private Sub HandleHelp2Click()
    MsgBox "Display Help Screen.", vbInformation
End Sub


