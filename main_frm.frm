VERSION 5.00
Begin VB.Form frm_main 
   Caption         =   "Main Form"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form10"
   Picture         =   "main_frm.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   0
      Picture         =   "main_frm.frx":30556
      ScaleHeight     =   1155
      ScaleWidth      =   15315
      TabIndex        =   14
      Top             =   0
      Width           =   15375
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":363BE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8520
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COMPLAINT DETAILS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":66914
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   4335
   End
   Begin VB.CommandButton cus_rep 
      Caption         =   "CUSTOMER REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":96E6A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton coneec_rep 
      Caption         =   "CONNECTIONS REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":C73C0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   4335
   End
   Begin VB.CommandButton dis_rep 
      Caption         =   "DISCONNECT REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":F7916
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   4335
   End
   Begin VB.CommandButton sale_rep 
      Caption         =   "SALE OF CARDS REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":127E6C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   4335
   End
   Begin VB.CommandButton comp_rep 
      Caption         =   "COMPLAINT REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":1583C2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   4335
   End
   Begin VB.CommandButton bill_rep 
      Caption         =   "BILL REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Picture         =   "main_frm.frx":188918
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "BILL DETAILS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":1B8E6E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AREA WISE CONNECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":1E93C4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SALE OF CARDS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":21991A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DISCONNECT CONNECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":249E70
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BOOKING NEW CONNECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":27A3C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   4335
   End
   Begin VB.CommandButton cus_link 
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Picture         =   "main_frm.frx":2AA91C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "TELE NETWORKING SYSTEM"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   1200
      Width           =   15375
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bill_rep_Click()
DataReport5.Show
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frm_comp_det.Show
End Sub

Private Sub Command2_Click()
frm_booking.Show
End Sub

Private Sub Command3_Click()
frm_disconnect.Show
End Sub

Private Sub Command4_Click()
frm_sale_cards.Show
End Sub

Private Sub Command5_Click()
frm_area_det.Show
End Sub

Private Sub Command6_Click()
frm_bill.Show
End Sub

Private Sub Command9_Click()
Form1.Show
End Sub

Private Sub comp_rep_Click()
DataReport6.Show
End Sub

Private Sub coneec_rep_Click()
DataReport2.Show
End Sub

Private Sub cus_link_Click()
frm_cus.Show
End Sub

Private Sub cus_rep_Click()
DataReport1.Show
End Sub

Private Sub dis_rep_Click()
DataReport4.Show
End Sub

Private Sub sale_rep_Click()
DataReport3.Show
End Sub
