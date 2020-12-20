VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3408
   ClientLeft      =   3456
   ClientTop       =   3240
   ClientWidth     =   4896
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2346.323
   ScaleMode       =   0  'User
   ScaleWidth      =   4592.788
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   1284
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.micronetsoft.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   324
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   2532
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Visit our web site:"
      ForeColor       =   &H00000000&
      Height          =   204
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      ForeColor       =   &H00000000&
      Height          =   204
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1932
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   450.273
      X2              =   4165.028
      Y1              =   669.198
      Y2              =   669.198
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   450.273
      X2              =   4165.028
      Y1              =   660.936
      Y2              =   660.936
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2001 MicroNet Software Technology"
      ForeColor       =   &H00000000&
      Height          =   204
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   3972
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Database"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   468
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2292
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C00000&
      Height          =   228
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    'center the AboutBox on the screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    'Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
End Sub

