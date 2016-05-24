VERSION 5.00
Begin VB.Form frm_log 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form5"
   Picture         =   "log.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   7440
      Picture         =   "log.frx":417F
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   3000
      Width           =   555
   End
   Begin VB.CommandButton cmd_ok 
      Default         =   -1  'True
      Height          =   495
      Left            =   6600
      Picture         =   "log.frx":4A49
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "OK"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox pswd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   7080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox uname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7080
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   7
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frm_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
If uname.Text = "Admin" And pswd.Text = "Admin" Then
frm_log.Hide
frm_main.Show
Else
MsgBox "Invalid Username & Password", vbInformation
Exit Sub
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
