VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsole 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CONSOLE"
   ClientHeight    =   6108
   ClientLeft      =   2064
   ClientTop       =   2040
   ClientWidth     =   8064
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleWidth      =   8064
   Begin MSComctlLib.Toolbar tbrMain2 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8064
      _ExtentX        =   14224
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
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
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   1452
      Left            =   360
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2760
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   2561
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Work with dbuser table"
      Height          =   1212
      Left            =   3840
      TabIndex        =   13
      Top             =   600
      Width           =   3972
      Begin VB.CommandButton cmdEdit 
         Height          =   252
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Edit Table"
         Top             =   360
         Width           =   372
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   252
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Erase Table"
         Top             =   720
         Width           =   372
      End
      Begin VB.CommandButton cmdView 
         Height          =   252
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "View Data"
         Top             =   720
         Width           =   372
      End
      Begin VB.CommandButton cmdInsert 
         Height          =   252
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Insert Data"
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label11 
         Caption         =   "Edit Data"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "Erase All Data"
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
         Left            =   2520
         TabIndex        =   16
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label Label9 
         Caption         =   "View Data"
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
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "Insert Data"
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
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   1092
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Work with database objects"
      Height          =   1212
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3372
      Begin VB.CommandButton cmdDrop 
         Height          =   252
         Left            =   480
         TabIndex        =   2
         ToolTipText     =   "Delete Table"
         Top             =   720
         Width           =   372
      End
      Begin VB.CommandButton cmdCreate 
         Height          =   252
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Create Table"
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label4 
         Caption         =   "Delete dbuser table"
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
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   1932
      End
      Begin VB.Label Label2 
         Caption         =   "Create dbuser table"
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
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   2052
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Determine a date range:"
      Height          =   1572
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   7572
      Begin VB.CommandButton cmdWEndg 
         Height          =   252
         Left            =   3480
         TabIndex        =   31
         ToolTipText     =   "Enter Week Ending Date"
         Top             =   720
         Width           =   372
      End
      Begin VB.CommandButton cmdClearDates 
         Height          =   252
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   372
      End
      Begin VB.TextBox txtSunday 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   2160
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   972
      End
      Begin VB.TextBox txtMonday 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   972
      End
      Begin VB.CommandButton cmdWEnding 
         Height          =   252
         Left            =   3480
         TabIndex        =   9
         ToolTipText     =   "Enter Week Ending Date"
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label14 
         Caption         =   "Enter W/E Date using calendar control."
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
         Left            =   3960
         TabIndex        =   32
         Top             =   720
         Width           =   3492
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "through"
         Height          =   252
         Left            =   1320
         TabIndex        =   30
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label12 
         Caption         =   "Clear Dates"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Sunday"
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
         Left            =   2160
         TabIndex        =   25
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Monday "
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Enter W/E Date using text."
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
         Left            =   3960
         TabIndex        =   19
         Top             =   360
         Width           =   3132
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Access Objects - data from dbuser table"
      Height          =   2412
      Left            =   240
      TabIndex        =   22
      Top             =   1920
      Width           =   7572
      Begin VB.CommandButton cmdClearTable 
         Height          =   252
         Left            =   2400
         TabIndex        =   8
         ToolTipText     =   "Clear Table"
         Top             =   360
         Width           =   372
      End
      Begin VB.CommandButton cmdViewTable 
         Height          =   252
         Left            =   480
         TabIndex        =   7
         ToolTipText     =   "Load data"
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label7 
         Caption         =   "Clear Grid"
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
         Left            =   2880
         TabIndex        =   27
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Load Grid"
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
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   1212
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   6480
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
            Picture         =   "frmConsole.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":1AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsole.frx":1DF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7920
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ErrNumber As String
Dim ErrSource As String
Dim ErrDescription As String
Dim AccessConnect As String
Dim SQLServerConnect As String

Private Sub cmdClearDates_Click()
    'clear dates
    If txtMonday.Text <> "" And txtSunday.Text <> "" Then
        Msg = "This will clear the dates."
        Style = vbCritical + vbOK + vbDefaultButton2
        Ans = MsgBox(Msg, Style, "Setup")
        'ans will return as IDCANCEL = 2, IDOK = 1
        If Ans = IDOK Then
            txtMonday.Text = ""
            txtSunday.Text = ""
        End If
    End If
End Sub

Private Sub cmdClearTable_Click()
    'clear table
    If FlexGrid1.Rows > 2 Then
        Msg = "This will clear the grid."
        Style = vbCritical + vbOK + vbDefaultButton2
        Ans = MsgBox(Msg, Style, "Setup")
        'ans will return as IDCANCEL = 2, IDOK = 1
        If Ans = IDOK Then
            ClearFlexGrid
        End If
    End If
End Sub

Private Sub cmdCreate_Click()
    If g_oTable.TableExist Then
        Msg = "This will delete existing table."
        Style = vbCritical + vbOK + vbDefaultButton2
        Ans = MsgBox(Msg, Style, "Setup")
        'ans will return as IDCANCEL = 2, IDOK = 1
        If Ans = IDOK Then
            g_oTable.Create
            MsgBox "DbUser - table has been created.", vbInformation
        End If
    Else
        g_oTable.Create
        MsgBox "DbUser - table has been created.", vbInformation
    End If
End Sub

Private Sub cmdDelete_Click()
    'delete entries in dbuser table
    If g_oTable.TableExist Then
        g_oTable.Delete
        MsgBox "DbUser - table has been cleared.", vbInformation
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
End Sub

Private Sub cmdDrop_Click()
    'drop dbuser table
    If g_oTable.TableExist Then
        Msg = "This will delete existing table."
        Style = vbCritical + vbOK + vbDefaultButton2
        Ans = MsgBox(Msg, Style, "Setup")
        'ans will return as IDCANCEL = 2, IDOK = 1
        If Ans = IDOK Then
            g_oTable.Drop
            MsgBox "DbUser - table has been deleted.", vbInformation
        End If
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
End Sub

Private Sub cmdEdit_Click()
    If g_oTable.TableExist Then
        frmEditDb.Show vbModal
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload frmConsole
End Sub

Private Sub cmdInsert_Click()
    If g_oTable.TableExist Then
        If g_oTable.IsEmptyTable Then
            g_oTable.Insert
            MsgBox "DbUser - data has been inserted.", vbInformation
        Else
            Msg = "Table already contains data."
            MsgBox Msg, 64, "Error"
        End If
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
End Sub
Private Sub cmdView_Click()
    If g_oTable.TableExist Then
        frmDisplay.Show vbModal
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
End Sub

Private Sub cmdViewTable_Click()
    
    Dim sDatabase As Database
    Dim sConnect As String
    Dim sDsn As String
    Dim tempstr As String
    Dim tServer As String
    Dim tDB As String
    Dim tId As String
    Dim tPwd As String
    
    Dim i As Integer
    Dim j As Long
    Dim sSQL As String
    Dim rs As New CSQLSelect

    On Error GoTo ERROR_run
    
    If g_oTable.TableExist Then
    
        rs.Initialize g_dsPA

        rs.setSql "SELECT uid, name, status FROM dbuser ORDER BY uid"

        If rs.execute = False Then
            MsgBox rs.getError, vbOKOnly, "SQL Error"
            'GoTo Error_run
        End If
                
        'Add each record until the end of the file
        j = 1
        If (Not rs.getEOF) Then
            If FlexGrid1.Rows <= 2 Then
                'FlexGrid1, name of flexgrid
                'DBGrid1, name of dbgrid
                With FlexGrid1
                    .Cols = 4
                    .ColWidth(0) = 250
                    .ColWidth(1) = 1000
                    .ColWidth(2) = 2500
                    .ColWidth(3) = 720
                    .Row = 0
                    .Col = 0
                    .Text = ""
                    .Col = 1
                    .Text = "UID"
                    .Col = 2
                    .Text = "Name"
                    .Col = 3
                    .Text = "Status"
                End With
            
                Do Until rs.getEOF
                    FlexGrid1.Row = j
                    For i = 1 To 3
                        'MsgBox rs.getRsValue(i-1)
                        With FlexGrid1
                            .Col = i
                            .Text = rs.getRsValue(i - 1)
                        End With
                    Next
                    ' Get next row and add another row to the grid
                    rs.moveNext
                    If Not rs.getEOF Then
                        j = j + 1
                        FlexGrid1.Rows = j + 1
                    End If
                Loop
            End If  'if flexgrdi1.rows <= 2
        Else
            Msg = "Table is empty."
            MsgBox Msg, 16, "Error"
        End If
    Else
        Msg = "Table not found."
        MsgBox Msg, 16, "Error"
    End If
    
    Exit Sub
    
ERROR_run:
    MsgBox "Please Contact Administrator!", vbCritical
End Sub

Private Sub cmdWEndg_Click()
    gStartDate = ""
    gEndDate = ""
    frmWEndg.Show vbModal
    If gStartDate <> "" And gEndDate <> "" Then
        txtMonday = gStartDate
        txtSunday = gEndDate
    End If
End Sub

Private Sub cmdWEnding_Click()
    gStartDate = ""
    gEndDate = ""
    frmWEnding.Show vbModal
    If gStartDate <> "" And gEndDate <> "" Then
        txtMonday = gStartDate
        txtSunday = gEndDate
    End If
End Sub

Private Sub Form_Load()
    'center form on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    ' the \ operator will return an integer value.
    ' the / operator will return a float value.
    gDbuser = False
    txtMonday.Text = ""
    txtSunday.Text = ""
    InitFlexGrid
End Sub

Private Sub tbrMain2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ksave is the key name of the button
    Select Case Button.Key
        Case "kFinish2"
            HandleFinish2Click
        Case "kCancel2"
            'clear button
            HandleCancel2Click
        Case "kHelp2"
            HandleHelp2Click
    End Select

End Sub
Private Sub HandleFinish2Click()
    'MsgBox "Finish.", vbInformation
    Unload frmConsole
End Sub

Private Sub HandleCancel2Click()
    Unload frmConsole
End Sub

Private Sub HandleHelp2Click()
    MsgBox "Display Help Screen.", vbInformation
End Sub

Public Sub ClearFlexGrid()
    'clear items in flexgrid
    Dim flag As Integer
    
    If FlexGrid1.Rows > 2 Then
        flag = True
        Do While flag
            'MsgBox FlexGrid1.Rows
            If FlexGrid1.Rows <= 2 Then
                flag = False
                FlexGrid1.Clear
                InitFlexGrid
            Else
                FlexGrid1.RemoveItem (1)
            End If
        Loop
        FlexGrid1.Refresh
    End If

End Sub

Public Sub InitFlexGrid()
    'initialize the flexgrid
    With FlexGrid1
        .Cols = 4
        .ColWidth(0) = 250
        .ColWidth(1) = 1000
        .ColWidth(2) = 2500
        .ColWidth(3) = 720
        .Row = 0
        .Col = 0
        .Text = ""
        .Col = 1
        .Text = ""
        .Col = 2
        .Text = ""
        .Col = 3
        .Text = ""
    End With
End Sub
