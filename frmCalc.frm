VERSION 5.00
Begin VB.Form frmCalc 
   Caption         =   "Calculator"
   ClientHeight    =   2532
   ClientLeft      =   3900
   ClientTop       =   2724
   ClientWidth     =   3312
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2532
   ScaleWidth      =   3312
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   363
      Left            =   1350
      Picture         =   "frmCalc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "Previous"
      Top             =   1980
      Width           =   454
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1530
      Width           =   444
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1530
      Width           =   444
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1530
      Width           =   444
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   444
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   444
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1350
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   444
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   444
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   444
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1350
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   630
      Width           =   444
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1980
      Width           =   444
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1980
      Width           =   444
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   18
      Top             =   1980
      Width           =   984
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   444
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1530
      Width           =   444
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1530
      Width           =   444
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2610
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   444
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   0
      Top             =   630
      Width           =   444
   End
   Begin VB.CommandButton cmdClearEntry 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2610
      TabIndex        =   1
      Top             =   630
      Width           =   444
   End
   Begin VB.Label lblReadOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   19
      Top             =   96
      Width           =   3036
   End
   Begin VB.Label lblReadout2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2976
   End
   Begin VB.Shape Box28 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   1890
      Left            =   1980
      Top             =   540
      Width           =   1170
   End
   Begin VB.Shape Box27 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   1890
      Left            =   180
      Top             =   540
      Width           =   1710
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastInput As Integer
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim Op1 As Double
Dim Op2 As Double
Dim OpFlag As String

Const STATE_NONE = 0
Const STATE_NUMS = 1
Const STATE_OPS = 2
Const STATE_CE = 3

Const KEY_LBUTTON = &H1
Const KEY_RBUTTON = &H2
Const KEY_CANCEL = &H3
Const KEY_MBUTTON = &H4         ' NOT contiguous with L & R BUTTON
Const KEY_BACK = &H8
Const KEY_TAB = &H9
Const KEY_CLEAR = &HC
Const KEY_RETURN = &HD
Const KEY_SHIFT = &H10
Const KEY_CONTROL = &H11
Const KEY_MENU = &H12
Const KEY_PAUSE = &H13
Const KEY_CAPITAL = &H14
Const KEY_ESCAPE = &H1B
Const KEY_SPACE = &H20
Const KEY_PRIOR = &H21
Const KEY_NEXT = &H22
Const KEY_END = &H23
Const KEY_HOME = &H24
Const KEY_LEFT = &H25
Const KEY_UP = &H26
Const KEY_RIGHT = &H27
Const KEY_DOWN = &H28
Const KEY_SELECT = &H29
Const KEY_PRINT = &H2A
Const KEY_EXECUTE = &H2B
Const KEY_SNAPSHOT = &H2C
Const KEY_INSERT = &H2D
Const KEY_DELETE = &H2E
Const KEY_HELP = &H2F

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

Const KEY_NUMPAD0 = &H60
Const KEY_NUMPAD1 = &H61
Const KEY_NUMPAD2 = &H62
Const KEY_NUMPAD3 = &H63
Const KEY_NUMPAD4 = &H64
Const KEY_NUMPAD5 = &H65
Const KEY_NUMPAD6 = &H66
Const KEY_NUMPAD7 = &H67
Const KEY_NUMPAD8 = &H68
Const KEY_NUMPAD9 = &H69
Const KEY_MULTIPLY = &H6A
Const KEY_ADD = &H6B
Const KEY_SEPARATOR = &H6C
Const KEY_SUBTRACT = &H6D
Const KEY_DECIMAL = &H6E
Const KEY_DIVIDE = &H6F
Const KEY_F1 = &H70
Const KEY_F2 = &H71
Const KEY_F3 = &H72
Const KEY_F4 = &H73
Const KEY_F5 = &H74
Const KEY_F6 = &H75
Const KEY_F7 = &H76
Const KEY_F8 = &H77
Const KEY_F9 = &H78
Const KEY_F10 = &H79
Const KEY_F11 = &H7A
Const KEY_F12 = &H7B
Const KEY_F13 = &H7C
Const KEY_F14 = &H7D
Const KEY_F15 = &H7E
Const KEY_F16 = &H7F
Const KEY_NUMLOCK = &H90
Const KEY_PERIOD = &HBE
Const KEY_EQUAL = &HBB
Const KEY_MINUS = &HBD
Const KEY_SLASH = &HBF
Const KEY_C = &H43  ' C
Const SHIFT_MASK = 1

'Private msOpenArgs As String

Private Sub cmd0_Click()
    HandleNumberClick "0"
End Sub

Private Sub cmd0_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd1_Click()
    HandleNumberClick "1"
End Sub

Private Sub cmd1_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd2_Click()
    HandleNumberClick "2"
End Sub

Private Sub cmd2_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd3_Click()
    HandleNumberClick "3"
End Sub

Private Sub cmd3_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd4_Click()
    HandleNumberClick "4"
End Sub

Private Sub cmd4_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd5_Click()
    HandleNumberClick "5"
End Sub

Private Sub cmd5_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd6_Click()
    HandleNumberClick "6"
End Sub

Private Sub cmd6_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd7_Click()
    HandleNumberClick "7"
End Sub

Private Sub cmd7_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd8_Click()
    HandleNumberClick "8"
End Sub

Private Sub cmd8_keydown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmd9_Click()
    HandleNumberClick "9"
End Sub

Private Sub cmd9_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdAdd_Click()
    HandleOperatorClick "+"
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdBack_Click()
    HandleBackspace
End Sub

Private Sub cmdBack_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdCancel_Click()
'    Unload frmCalc        'Screen.ActiveForm
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdClear_Click()
    HandleClearClick
End Sub

Private Sub cmdClear_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdClearEntry_Click()
    HandleClearEntry
End Sub

Private Sub cmdClearEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdDecimal_Click()
    HandleDecimalClick
End Sub

Private Sub cmdDecimal_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdDivide_Click()
    HandleOperatorClick "/"
End Sub

Private Sub cmdDivide_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdEquals_Click()
    HandleOperatorClick "="
End Sub

Private Sub cmdEquals_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdMultiply_Click()
    HandleOperatorClick "*"
End Sub

Private Sub cmdMultiply_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdOK_Click()
'    HandleOperatorClick "="
'    DoEvents
'    Unload frmCalc
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub

Private Sub cmdSubtract_Click()
    HandleOperatorClick "-"
End Sub

Private Sub cmdSubtract_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleNumberPress KeyCode, Shift
End Sub
Private Sub Form_Load()
    Dim strArgs As String
    Dim intPos As Integer

    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    lblReadout2.Top = lblReadOut.Top
    lblReadout2.Left = lblReadOut.Left
    InitCalc
End Sub

Private Sub HandleBackspace()
    Dim varCaption As Variant

    varCaption = lblReadOut.Caption
    If Not IsNull(varCaption) And varCaption <> Empty Then
        lblReadOut.Caption = Left(varCaption, Len(varCaption) - 1)
        SetReadout
    End If
    Screen.ActiveForm.Refresh
    If IsNull(lblReadOut.Caption) Then
        LastInput = STATE_NONE
        DecimalFlag = False
    Else
        LastInput = STATE_NUMS
        DecimalFlag = (InStr(lblReadOut.Caption, ".") > 0)
    End If
End Sub

Private Sub HandleClearClick()
    lblReadOut.Caption = ""
    SetReadout
    InitCalc
End Sub

Private Sub HandleClearEntry()
    lblReadOut.Caption = ""
    SetReadout
    DecimalFlag = False
    LastInput = STATE_CE
End Sub

Private Sub HandleDecimalClick()
    If LastInput <> STATE_NUMS Then
        lblReadOut.Caption = "0."
        SetReadout
    ElseIf Not DecimalFlag Then
        lblReadOut.Caption = lblReadOut.Caption & "."
        SetReadout
    End If
    DecimalFlag = True
    LastInput = STATE_NUMS
End Sub

Private Sub HandleNumberClick(strNum As String)
    Dim varCaption As Variant
    Dim varLen As Variant

    If LastInput <> STATE_NUMS Then
        lblReadOut.Caption = ""
        SetReadout
        DecimalFlag = False
    End If

    varCaption = lblReadOut.Caption
    varLen = Len(varCaption)
    lblReadOut.Caption = lblReadOut.Caption & strNum
    SetReadout
    If varLen > 0 And Left(varCaption, 1) = "0" And InStr(varCaption, ".") = 0 Then
        lblReadOut.Caption = Mid(lblReadOut.Caption, 2)
        SetReadout
    End If
    LastInput = STATE_NUMS
End Sub

Private Sub HandleNumberPress(KeyCode As Integer, Shift As Integer)
    Dim strChar As String
    Dim intShiftDown

    intShiftDown = (Shift And SHIFT_MASK)
    Select Case KeyCode
        ' 0 - 9
        Case 48 To 57
            If intShiftDown Then
                Select Case KeyCode
                    ' Handle Shift-8 ("*")
                    Case 48 + 8
                        Screen.ActiveForm.Controls("cmdMultiply").SetFocus
                        HandleOperatorClick "*"
                End Select
            Else
                strChar = Chr$(KeyCode)
                Screen.ActiveForm.Controls("cmd" & strChar).SetFocus
                HandleNumberClick strChar
            End If
            ' Backspace
        Case KEY_BACK
            Screen.ActiveForm.Controls("cmdBack").SetFocus
            HandleBackspace
        Case KEY_DELETE
            Screen.ActiveForm.Controls("cmdClearEntry").SetFocus
            HandleClearEntry
            ' Numpad 0 - Numpad 9
        Case KEY_NUMPAD0 To KEY_NUMPAD9
            strChar = Chr$(KeyCode - KEY_NUMPAD0 + Asc("0"))
            Screen.ActiveForm.Controls("cmd" & strChar).SetFocus
            HandleNumberClick strChar
            ' Period and Decimal
        Case KEY_PERIOD, KEY_DECIMAL
            Screen.ActiveForm.Controls("cmdDecimal").SetFocus
            HandleDecimalClick
        Case KEY_SUBTRACT, KEY_MINUS
            Screen.ActiveForm.Controls("cmdSubtract").SetFocus
            HandleOperatorClick "-"
        Case KEY_MULTIPLY
            Screen.ActiveForm.Controls("cmdMultiply").SetFocus
            HandleOperatorClick "*"
        Case KEY_ADD
            Screen.ActiveForm.Controls("cmdAdd").SetFocus
            HandleOperatorClick "+"
        Case KEY_DIVIDE, KEY_SLASH
            Screen.ActiveForm.Controls("cmdDivide").SetFocus
            HandleOperatorClick "/"
        Case KEY_EQUAL, KEY_RETURN
            If intShiftDown Then
                Screen.ActiveForm.Controls("cmdAdd").SetFocus
                HandleOperatorClick "+"
            Else
                Screen.ActiveForm.Controls("cmdEquals").SetFocus
                HandleOperatorClick "="
                If KeyCode = KEY_RETURN Then
                    DoEvents
                    Me.Visible = False
                End If
            End If
        Case KEY_C
            Screen.ActiveForm.Controls("cmdClear").SetFocus
            HandleClearClick
    End Select
End Sub

Private Sub HandleOperatorClick(strOp As String)
    If LastInput = STATE_NUMS Then
        NumOps = NumOps + 1
    End If
    If NumOps = 1 Then
        Op1 = Val(lblReadOut.Caption)
    ElseIf NumOps = 2 Then
        Op2 = Val(lblReadOut.Caption)
        Select Case OpFlag
            Case "+"
                Op1 = Op1 + Op2
            Case "-"
                Op1 = Op1 - Op2
            Case "*"
                Op1 = Op1 * Op2
            Case "/"
                If Op2 = 0 Then
                    MsgBox "Can't Divide by Zero", 48, "Calculator"
                Else
                    Op1 = Op1 / Op2
                End If
            Case "="
                Op1 = Op2
        End Select
        lblReadOut.Caption = Format$(Op1)
        SetReadout
        NumOps = 1
    End If
    LastInput = STATE_OPS
    OpFlag = strOp
End Sub

Private Sub InitCalc()
    lblReadOut.Caption = ""
    SetReadout
    LastInput = STATE_NONE
    DecimalFlag = False
    NumOps = 0
End Sub

Private Sub SetReadout()
    Dim varReadout As Variant
    Dim dblReadout As Double

    varReadout = lblReadOut.Caption
    If IsNull(varReadout) Then varReadout = "0"
    dblReadout = Val(varReadout)
    lblReadout2.Caption = Format(dblReadout, "###,###,###,###,##0.00#####")
End Sub
