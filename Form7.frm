VERSION 5.00
Begin VB.Form frm_sale_cards 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sales Details"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form7"
   ScaleHeight     =   4725
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   6240
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   480
      Width           =   3255
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.TextBox amt 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   6120
      TabIndex        =   12
      Top             =   2640
      Width           =   3255
      Begin VB.CommandButton cmd_clr 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Picture         =   "Form7.frx":19C2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Picture         =   "Form7.frx":206D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Save 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Picture         =   "Form7.frx":2718
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Of Cards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   5415
      Begin VB.CheckBox chk5_san 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SANCHARNET CARD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chk4_call 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CALL NOW CARD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox chk6_itc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ITC "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chk3_top 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOP UP CARD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chk2_rec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECHARGE COUPAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chk1_sim 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SIM CARD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox cus_name 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cus_ref 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "TELECOMMUNICATION SYSTEM OF NETWORKS"
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
      TabIndex        =   4
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Ref"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frm_sale_cards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmd_clr_Click()
cus_ref.Text = ""
cus_name.Text = ""
amt.Text = ""
chk1_sim.Value = 0
chk2_rec.Value = 0
chk3_top.Value = 0
chk4_call.Value = 0
chk5_san.Value = 0
chk6_itc.Value = 0

If rs.RecordCount = 0 Then
Else
rs.MoveLast
cus_ref.Text = rs(0) + 1
End If

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()

If cus_ref.Text = "" Or cus_name.Text = "" Or amt.Text = "" Then
MsgBox "Empty Fields"
Else
rs.AddNew
rs.Fields(0) = cus_ref.Text
rs.Fields(1) = cus_name.Text
rs.Fields(2) = amt.Text

If chk1_sim.Value = 1 Then
rs.Fields(3) = chk1_sim.Caption
End If

If chk2_rec.Value = 1 Then
rs.Fields(4) = chk2_rec.Caption
End If

If chk3_top.Value = 1 Then
rs.Fields(5) = chk3_top.Caption
End If

If chk4_call.Value = 1 Then
rs.Fields(6) = chk4_call.Caption
End If

If chk5_san.Value = 1 Then
rs.Fields(7) = chk5_san.Caption
End If

If chk6_itc.Value = 1 Then
rs.Fields(8) = chk6_itc.Caption
End If

rs.Update
MsgBox "Saved"

cus_ref.Text = ""
cus_name.Text = ""
amt.Text = ""
chk1_sim.Value = 0
chk2_rec.Value = 0
chk3_top.Value = 0
chk4_call.Value = 0
chk5_san.Value = 0
chk6_itc.Value = 0

rs.Close
rs.Open "select * from sale_of_cards", cn, adOpenDynamic, adLockPessimistic

End If

If rs.RecordCount = 0 Then
Else
rs.MoveLast
cus_ref.Text = rs(0) + 1
End If

End Sub

Private Sub Form_Load()

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from sale_of_cards", cn, adOpenDynamic, adLockPessimistic

If rs.RecordCount = 0 Then
Else
rs.MoveLast
cus_ref.Text = rs(0) + 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
cn.Close
End Sub
