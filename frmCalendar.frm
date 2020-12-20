VERSION 5.00
Begin VB.Form frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   2844
   ClientLeft      =   4416
   ClientTop       =   3564
   ClientWidth     =   3084
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
   ScaleHeight     =   2844
   ScaleWidth      =   3084
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdNextYear 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2700
      Picture         =   "frmCalendar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Yr >>"
      Top             =   135
      Width           =   245
   End
   Begin VB.CommandButton CmdPreviousYear 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      Picture         =   "frmCalendar.frx":00DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "<< Yr"
      Top             =   135
      Width           =   245
   End
   Begin VB.CommandButton CmdNextMonth 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1460
      Picture         =   "frmCalendar.frx":01F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Mnth >>"
      Top             =   135
      Width           =   245
   End
   Begin VB.CommandButton CmdPreviousMonth 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      Picture         =   "frmCalendar.frx":02CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "<< Mnth"
      Top             =   135
      Width           =   245
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1440
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1020
   End
   Begin VB.TextBox Day 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.TextBox Year 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   585
   End
   Begin VB.TextBox Month 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      Height          =   1932
      Left            =   225
      Top             =   480
      Width           =   2600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   252
      Left            =   700
      TabIndex        =   60
      Top             =   2520
      Width           =   612
   End
   Begin VB.Label lblDay1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Su"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   228
      TabIndex        =   10
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   11
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   972
      TabIndex        =   12
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "We"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1344
      TabIndex        =   13
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1716
      TabIndex        =   14
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2088
      TabIndex        =   15
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lblDay7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2460
      TabIndex        =   16
      Top             =   612
      Width           =   372
   End
   Begin VB.Label lbl11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   228
      TabIndex        =   17
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   18
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   19
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1344
      TabIndex        =   20
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   21
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   22
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   23
      Top             =   936
      Width           =   360
   End
   Begin VB.Label lbl21 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   228
      TabIndex        =   24
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl22 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   25
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   26
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl24 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1344
      TabIndex        =   27
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   28
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   29
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl27 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   30
      Top             =   1152
      Width           =   360
   End
   Begin VB.Label lbl31 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   228
      TabIndex        =   31
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   32
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl33 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   33
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   1344
      TabIndex        =   34
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl35 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   35
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl36 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   36
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl37 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   37
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl41 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   240
      TabIndex        =   38
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl42 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   39
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl43 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   40
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl44 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1344
      TabIndex        =   41
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl45 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   42
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl46 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   43
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl47 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   44
      Top             =   1608
      Width           =   360
   End
   Begin VB.Label lbl51 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   228
      TabIndex        =   45
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl52 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   46
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl53 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   47
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl54 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1344
      TabIndex        =   48
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl55 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   49
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl56 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   50
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl57 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   51
      Top             =   1836
      Width           =   360
   End
   Begin VB.Label lbl61 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   228
      TabIndex        =   52
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl62 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   600
      TabIndex        =   53
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl63 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   972
      TabIndex        =   54
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl64 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1344
      TabIndex        =   55
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl65 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1716
      TabIndex        =   56
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl66 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2088
      TabIndex        =   57
      Top             =   2052
      Width           =   360
   End
   Begin VB.Label lbl67 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   2460
      TabIndex        =   58
      Top             =   2052
      Width           =   360
   End
   Begin VB.Line Line718 
      BorderColor     =   &H0080FFFF&
      X1              =   225
      X2              =   2825
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Set the first displayed day of the week.  In the
' US, this is Sunday (1).  In other countries,
' use the appropriate number (1 == Sunday, 7 == Saturday).
Const FIRST_DAY = 1

' Color to show weekend days.
Const COLOR_WEEKEND = vbRed     '255
Const COLOR_WEEKDAY = vbBlack   '0
Const COLOR_DAY = vbBlue        '16711680

Const D_SUN = "Su"
Const D_MON = "Mo"
Const D_TUE = "Tu"
Const D_WED = "We"
Const D_THU = "Th"
Const D_FRI = "Fr"
Const D_SAT = "Sa"

Dim astrDays(1 To 7) As String

Dim intStartDOW As Integer

' Store away today's date.
Dim intYearToday As Integer
Dim intMonthToday As Integer
Dim intDayToday As Integer

Dim aintMonthLen(1 To 12) As Integer
Dim strSelected As String
Dim mCol As Integer
Dim mRow As Integer
Dim mStr As String

' Constants used to control movement on the form.
' These constants match the interval values
' needed by DateAdd().
Const CHANGE_DAY = "d"
Const CHANGE_MONTH = "m"
Const CHANGE_YEAR = "yyyy"
Const CHANGE_WEEK = "ww"

Const MOVE_FORWARD = 0
Const MOVE_BACKWARD = 1

' Constant month values.
Const M_JAN = 1
Const M_FEB = 2
Const M_MAR = 3
Const M_APR = 4
Const M_MAY = 5
Const M_JUN = 6
Const M_JUL = 7
Const M_AUG = 8
Const M_SEP = 9
Const M_OCT = 10
Const M_NOV = 11
Const M_DEC = 12

' Key Codes
Const KEY_LBUTTON = &H1
Const KEY_RBUTTON = &H2
Const KEY_CANCEL = &H3
Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
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

' Shift parameter masks
Const SHIFT_MASK = 1
Const CTRL_MASK = 2
Const ALT_MASK = 4

Private Function Base7(wValue As Integer)
    ' Convert a number, up to 48 decimal, into base 7.
    Base7 = (wValue \ 7) & (wValue Mod 7)
End Function

Private Sub ChangeDate(strMoveUnit As String, intDirection As Integer)
    ' Called from OnPush property of the next/previous month/year buttons.
    Dim intMonth As Integer
    Dim intYear As Integer
    Dim intDay As Integer
    Dim varDate As Variant
    Dim varOldDate As Variant
    Dim intInc As Integer
    Dim rstrInterval As String

    On Error GoTo ChangeDateError

    ' Get the current values from the form.
    intYear = Me.Year
    intMonth = Me.Month
    intDay = Me.Day

    intInc = IIf(intDirection = MOVE_FORWARD, 1, -1)
    varOldDate = DateSerial(intYear, intMonth, intDay)
    varDate = DateAdd(strMoveUnit, intInc, varOldDate)


    If (intDirection = MOVE_BACKWARD And varDate > varOldDate) Then
        ' This should only happen when you go backward from
        ' 1/1/100 to 12/31/1999.
        Exit Sub
    End If

    intMonth = DatePart("m", varDate)
    intYear = DatePart("yyyy", varDate)
    Me.Day = DatePart("d", varDate)

    ' If the month and year haven't changed, then just
    ' move to the selected day.  It's a lot faster.
    If Me.Month = intMonth And Me.Year = intYear Then
        HandleIndent "lbl" & Day2Button((Me!Day), intStartDOW)
    Else
        ' Set the values on the form and then display the new calendar.
        Me.Month = intMonth
        Me.txtMonth = GetMonthName(intMonth)
        Me.Year = intYear
        DisplayCal
        txtDate.Text = Me.Month & "/" & Me.Day & "/" & Me.Year
    End If

ChangeDateExit:
    Exit Sub

ChangeDateError:
    Resume ChangeDateExit
End Sub

Private Sub cmdCancel_Click()
    Unload frmCalendar
End Sub

Private Sub CmdNextMonth_Click()
    ChangeDate CHANGE_MONTH, MOVE_FORWARD
End Sub

Private Sub CmdNextMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeys KeyCode, Shift
End Sub

Private Sub CmdNextYear_Click()
    ChangeDate CHANGE_YEAR, MOVE_FORWARD
End Sub

Private Sub CmdNextYear_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeys KeyCode, Shift
End Sub

Private Sub cmdOK_Click()
    ' Get the date that was chosen.
    Dim var As Variant
    var = SelectDate(strSelected)
End Sub

Private Sub CmdPreviousMonth_Click()
    ChangeDate CHANGE_MONTH, MOVE_BACKWARD
End Sub

Private Sub CmdPreviousMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeys KeyCode, Shift
End Sub

Private Sub CmdPreviousYear_Click()
    ChangeDate CHANGE_YEAR, MOVE_BACKWARD
End Sub

Private Sub CmdPreviousYear_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeys KeyCode, Shift
End Sub

Private Function Day2Button(wDay As Integer, intStartDay As Integer)
    Day2Button = Base7(wDay + intStartDay - 2 + 7) + 1
End Function

Private Function DaysInMonth(varMonthNumber As Variant) As Integer
    ' Get the number of days in the passed-in month.
    ' If the month isn't February, we know its length.
    If varMonthNumber <> M_FEB Then
        DaysInMonth = aintMonthLen(varMonthNumber)
    Else
        ' VB will know if it is a leap year
        ' Get the last day of the month of February for the currently displayed year.
        DaysInMonth = DatePart("d", DateSerial(Me!Year, M_MAR, 1) - 1)
    End If
End Function

Private Sub DisplayCal()
    ' Display the calendar.
    Static wInHere As Integer

    ' Eliminate a recursive call
    If wInHere Then Exit Sub
    wInHere = True

    ' Figure out the starting day of week for the given month.
    intStartDOW = FirstDOM((Me!Month), (Me!Year))
    
    ' Finally, display the calendar.
    ShowDate intStartDOW
    Me.Refresh
    wInHere = False
End Sub

Private Sub FillInStartValues()

    Dim varStartDate As Variant

    varStartDate = Date
    If IsNull(varStartDate) Or isEmpty(varStartDate) Then
        varStartDate = Date
    End If

    ' Store away the start date values
    Me.Month = DatePart("m", varStartDate)
    Me.Year = DatePart("yyyy", varStartDate)
    Me.Day = DatePart("d", varStartDate)
    Me.txtMonth = GetMonthName((Me!Month))

End Sub

Private Function FirstDOM(intMonth As Integer, intYear As Integer) As Integer
    ' Calculate the first day of the month in question.
    FirstDOM = DatePart("w", DateSerial(intYear, intMonth, 1), FIRST_DAY)
End Function

Private Sub FixDaysInMonth(intStartDay As Integer)
    ' Turn on and off buttons in the currently displayed month.
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intNumDays As Integer
    Dim intCount As Integer
    Dim strTemp As String


    intNumDays = DaysInMonth(Me!Month)
    ' If the chosen date is past the last day in this month,
    ' then just select the last day of this month.
    If Me.Day > intNumDays Then
        Me.Day = intNumDays
    End If

    intCount = 0
    For intRow = 1 To 6
        For intCol = 1 To 7
            If (intRow = 1) And (intCol < intStartDay) Then
                Me("lbl1" & intCol).Visible = False
            Else
                intCount = intCount + 1
                strTemp = "lbl" & intRow & intCol
                If intCount <= intNumDays Then
                    Me(strTemp).Visible = True
                    Me(strTemp).Caption = intCount
                Else
                    Me(strTemp).Visible = False
                End If
            End If
        Next intCol
    Next intRow
End Sub

Private Sub FixUpDisplay()
    ' Set the labels for the days of the week correctly,
    ' and set up the colors for the weekend days.

    Dim intCol As Integer
    Dim intRow As Integer
    Dim intLogicalDay As Integer
    Dim intDiff As Integer
    Dim ctl As Control

    intDiff = FIRST_DAY - 1

    For intCol = 1 To 7
        intLogicalDay = ((intCol + intDiff - 1) Mod 7) + 1
        Set ctl = Me("lblDay" & intCol)
        ctl.Caption = astrDays(intLogicalDay)
        If ((intLogicalDay - 1) Mod 6) = 0 Then
            ctl.ForeColor = COLOR_WEEKEND
            For intRow = 1 To 6
                Me("lbl" & intRow & intCol).ForeColor = COLOR_WEEKEND
            Next intRow
        End If

    Next intCol
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HandleKeys KeyCode, Shift
End Sub

Private Sub Form_Load()
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    ' Initialize the array of month lengths.
    aintMonthLen(M_JAN) = 31
    aintMonthLen(M_FEB) = 28       ' this may change
    aintMonthLen(M_MAR) = 31
    aintMonthLen(M_APR) = 30
    aintMonthLen(M_MAY) = 31
    aintMonthLen(M_JUN) = 30
    aintMonthLen(M_JUL) = 31
    aintMonthLen(M_AUG) = 31
    aintMonthLen(M_SEP) = 30
    aintMonthLen(M_OCT) = 31
    aintMonthLen(M_NOV) = 30
    aintMonthLen(M_DEC) = 31

    astrDays(1) = D_SUN
    astrDays(2) = D_MON
    astrDays(3) = D_TUE
    astrDays(4) = D_WED
    astrDays(5) = D_THU
    astrDays(6) = D_FRI
    astrDays(7) = D_SAT

    mCol = 0
    mRow = 0

    ' Get today's date
    intDayToday = DatePart("d", Date)
    intYearToday = DatePart("yyyy", Date)
    intMonthToday = DatePart("m", Date)
    
    'MsgBox intMonthToday & "/" & intDayToday & "/" & intYearToday

    ' Fill in the start values
    FillInStartValues

    ' Fix up the calendar display.
    FixUpDisplay

    ' Display the Calendar (which will get the month/year from the form)
    DisplayCal
    
    mStr = "lbl" & Day2Button((Me!Day), intStartDOW)
    Me(mStr).ForeColor = COLOR_DAY
    txtDate.Text = intMonthToday & "/" & intDayToday & "/" & intYearToday

End Sub

Private Function GetMonthName(intMonth As Integer) As String
    ' The year in the following expression is arbitrary.
    GetMonthName = Format(DateSerial(1995, intMonth, 1), "mmmm")
End Function

Private Sub HandleIndent(strNewSelect As String)

    If Len(strSelected) > 0 Then
        If strSelected <> strNewSelect Then
            GetColumn strSelected
            'this is for selected day
            If mCol = 1 Or mCol = 7 Then
                'either column is weekend
                Me(strSelected).ForeColor = COLOR_WEEKEND
            Else
                'column is during week
                Me(strSelected).ForeColor = COLOR_WEEKDAY
            End If
            Me(strNewSelect).ForeColor = COLOR_DAY
            'Me(strSelected).ForeColor = COLOR_WEEKDAY
        End If
    End If
    strSelected = strNewSelect
    Me.Day = Me(strSelected).Caption
    txtDate.Text = Me.Month & "/" & Me.Day & "/" & Me.Year
    
    'MsgBox "1 " & txtDate.Text
    gWkendDate = txtDate.Text

End Sub

Private Sub HandleKeys(KeyCode As Integer, Shift As Integer)

    ' Key Mappings:
    '
    ' Leftarrow = Previous Day
    ' Shift-Leftarrow = Previous Year
    ' Rightarrow = Next Day
    ' Shift-Rightarrow = Next Year
    ' Uparrow = Previous week
    ' Shift-Uparrow = Previous Month
    ' Dnarrow = Next Week
    ' Shift-Dnarrow = Next Month
    ' PgUp = Previous Month
    ' Shift-PgUp = Previous Year
    ' PgDn = Next Month
    ' Shift-PgDn = Next Year
    ' Home = Move to Today
    ' Shift-Home = Move to today in selected year.

    Dim ShiftDown As Integer

    ShiftDown = ((Shift And SHIFT_MASK) > 0)

    Select Case KeyCode
        Case KEY_ESCAPE
            Unload frmCalendar
        Case KEY_RETURN
            Me.Visible = False
        Case KEY_HOME
            If ShiftDown Then
                ' Use the selected year.
                MoveToToday False
            Else
                ' Use the actual current year.
                MoveToToday True
            End If
        Case KEY_PRIOR
            If ShiftDown Then
                ChangeDate CHANGE_YEAR, MOVE_BACKWARD
            Else
                ChangeDate CHANGE_MONTH, MOVE_BACKWARD
            End If
        Case KEY_NEXT
            If ShiftDown Then
                ChangeDate CHANGE_YEAR, MOVE_FORWARD
            Else
                ChangeDate CHANGE_MONTH, MOVE_FORWARD
            End If
        Case KEY_RIGHT
            If ShiftDown Then
                ' Move to next year
                ChangeDate CHANGE_YEAR, MOVE_FORWARD
            Else
                ChangeDate CHANGE_DAY, MOVE_FORWARD
            End If
        Case KEY_LEFT
            If ShiftDown Then
                ' Move to previous year
                ChangeDate CHANGE_YEAR, MOVE_BACKWARD
            Else
                ChangeDate CHANGE_DAY, MOVE_BACKWARD
            End If
        Case KEY_UP
            If ShiftDown Then
                ' Move to previous month
                ChangeDate CHANGE_MONTH, MOVE_BACKWARD
            Else
                ChangeDate CHANGE_WEEK, MOVE_BACKWARD
            End If
        Case KEY_DOWN
            If ShiftDown Then
                ' Move to next month
                ChangeDate CHANGE_MONTH, MOVE_FORWARD
            Else
                ChangeDate CHANGE_WEEK, MOVE_FORWARD
            End If
    End Select
    ' disregard the key press.
    KeyCode = 0
End Sub

Private Function HandleSelected(strName As String)
    HandleIndent strName
End Function

Private Sub MoveToToday(fUseCurrentYear As Integer)
    ' Month and year get filled in from the form.
    ' Go to the stored current date.
    Me.Month = intMonthToday
    Me.txtMonth = GetMonthName((Me!Month))
    Me.Day = intDayToday
    If fUseCurrentYear Then
        Me.Year = intYearToday
    End If

    ' Display the calendar.
    DisplayCal
End Sub

Private Function SelectDate(strName As String)
    HandleIndent strName
    Unload Me
End Function

Private Sub ShowDate(intStartDay As Integer)
    Dim newSelected As String

    ' Fix up the visible day buttons.
    FixDaysInMonth intStartDay

    ' Set the right button as depressed when the month is displayed.
    newSelected = "lbl" & Day2Button((Me!Day), intStartDay)
    HandleIndent newSelected
    Me.Refresh
End Sub

Private Sub lbl11_Click()
    Call HandleSelected("lbl11")
    Call SelectDate("lbl11")
End Sub

Private Sub lbl12_Click()
    Call HandleSelected("lbl12")
    Call SelectDate("lbl12")
End Sub

Private Sub lbl13_Click()
    Call HandleSelected("lbl13")
    Call SelectDate("lbl13")
End Sub

Private Sub lbl14_Click()
    Call HandleSelected("lbl14")
    Call SelectDate("lbl14")
End Sub

Private Sub lbl15_Click()
    Call HandleSelected("lbl15")
    Call SelectDate("lbl15")
End Sub

Private Sub lbl16_Click()
    Call HandleSelected("lbl16")
    Call SelectDate("lbl16")
End Sub

Private Sub lbl17_Click()
    Call HandleSelected("lbl17")
    Call SelectDate("lbl17")
End Sub

Private Sub lbl21_Click()
    Call HandleSelected("lbl21")
    Call SelectDate("lbl21")
End Sub

Private Sub lbl22_Click()
    Call HandleSelected("lbl22")
    Call SelectDate("lbl22")
End Sub

Private Sub lbl23_Click()
    Call HandleSelected("lbl23")
    Call SelectDate("lbl23")
End Sub

Private Sub lbl24_Click()
    Call HandleSelected("lbl24")
    Call SelectDate("lbl24")
End Sub

Private Sub lbl25_Click()
    Call HandleSelected("lbl25")
    Call SelectDate("lbl25")
End Sub

Private Sub lbl26_Click()
    Call HandleSelected("lbl26")
    Call SelectDate("lbl26")
End Sub

Private Sub lbl27_Click()
    Call HandleSelected("lbl27")
    Call SelectDate("lbl27")
End Sub

Private Sub lbl31_Click()
    Call HandleSelected("lbl31")
    Call SelectDate("lbl31")
End Sub

Private Sub lbl32_Click()
    Call HandleSelected("lbl32")
    Call SelectDate("lbl32")
End Sub

Private Sub lbl33_Click()
    Call HandleSelected("lbl33")
    Call SelectDate("lbl33")
End Sub

Private Sub lbl34_Click()
    Call HandleSelected("lbl34")
    Call SelectDate("lbl34")
End Sub

Private Sub lbl35_Click()
    Call HandleSelected("lbl35")
    Call SelectDate("lbl35")
End Sub

Private Sub lbl36_Click()
    Call HandleSelected("lbl36")
    Call SelectDate("lbl36")
End Sub

Private Sub lbl37_Click()
    Call HandleSelected("lbl37")
    Call SelectDate("lbl37")
End Sub

Private Sub lbl41_Click()
    Call HandleSelected("lbl41")
    Call SelectDate("lbl41")
End Sub

Private Sub lbl42_Click()
    Call HandleSelected("lbl42")
    Call SelectDate("lbl42")
End Sub

Private Sub lbl43_Click()
    Call HandleSelected("lbl43")
    Call SelectDate("lbl43")
End Sub

Private Sub lbl44_Click()
    Call HandleSelected("lbl44")
    Call SelectDate("lbl44")
End Sub

Private Sub lbl45_Click()
    Call HandleSelected("lbl45")
    Call SelectDate("lbl45")
End Sub

Private Sub lbl46_Click()
    Call HandleSelected("lbl46")
    Call SelectDate("lbl46")
End Sub

Private Sub lbl47_Click()
    Call HandleSelected("lbl47")
    Call SelectDate("lbl47")
End Sub

Private Sub lbl51_Click()
    Call HandleSelected("lbl51")
    Call SelectDate("lbl51")
End Sub

Private Sub lbl52_Click()
    Call HandleSelected("lbl52")
    Call SelectDate("lbl52")
End Sub

Private Sub lbl53_Click()
    Call HandleSelected("lbl53")
    Call SelectDate("lbl53")
End Sub

Private Sub lbl54_Click()
    Call HandleSelected("lbl54")
    Call SelectDate("lbl54")
End Sub

Private Sub lbl55_Click()
    Call HandleSelected("lbl55")
    Call SelectDate("lbl55")
End Sub

Private Sub lbl56_Click()
    Call HandleSelected("lbl56")
    Call SelectDate("lbl56")
End Sub

Private Sub lbl57_Click()
    Call HandleSelected("lbl57")
    Call SelectDate("lbl57")
End Sub

Private Sub lbl61_Click()
    Call HandleSelected("lbl61")
    Call SelectDate("lbl61")
End Sub

Private Sub lbl62_Click()
    Call HandleSelected("lbl62")
    Call SelectDate("lbl62")
End Sub

Private Sub lbl63_Click()
    Call HandleSelected("lbl63")
    Call SelectDate("lbl63")
End Sub

Private Sub lbl64_Click()
    Call HandleSelected("lbl64")
    Call SelectDate("lbl64")
End Sub

Private Sub lbl65_Click()
    Call HandleSelected("lbl65")
    Call SelectDate("lbl65")
End Sub

Private Sub lbl66_Click()
    Call HandleSelected("lbl66")
    Call SelectDate("lbl66")
End Sub

Private Sub lbl67_Click()
    Call HandleSelected("lbl67")
    Call SelectDate("lbl67")
End Sub

Public Sub GetColumn(tstr As String)
    'strip the label and get column number
    mCol = Mid(tstr, 5, 1)
End Sub
