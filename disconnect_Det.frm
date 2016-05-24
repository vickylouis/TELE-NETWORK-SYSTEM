VERSION 5.00
Begin VB.Form frm_disconnect 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cancellatin Details"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form9"
   ScaleHeight     =   5820
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
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
      Left            =   5520
      Picture         =   "disconnect_Det.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5160
      Width           =   1335
   End
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
      Left            =   4080
      Picture         =   "disconnect_Det.frx":06AB
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd_save 
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
      Left            =   2640
      Picture         =   "disconnect_Det.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox ctime 
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox cdt 
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   6000
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox ph_no 
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
      Left            =   2760
      TabIndex        =   13
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Connection"
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
      Height          =   3375
      Left            =   5400
      TabIndex        =   10
      Top             =   1440
      Width           =   4935
      Begin VB.ComboBox tr_from 
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
         ItemData        =   "disconnect_Det.frx":1401
         Left            =   2400
         List            =   "disconnect_Det.frx":1414
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox tr_to 
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
         ItemData        =   "disconnect_Det.frx":1457
         Left            =   2400
         List            =   "disconnect_Det.frx":146A
         TabIndex        =   20
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox tel_no 
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
         Left            =   2400
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox new_addr 
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
         Height          =   765
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transfer From"
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
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transfer To"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Tel Number"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Address"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.ComboBox can_type 
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
      ItemData        =   "disconnect_Det.frx":14AD
      Left            =   2760
      List            =   "disconnect_Det.frx":14B7
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox addr 
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
      Height          =   765
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cus_code 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
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
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancellation Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cust_code"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
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
      Left            =   360
      TabIndex        =   4
      Top             =   2040
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
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TelePhone Number"
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
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
End
Attribute VB_Name = "frm_disconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub can_type_Click()
If can_type.Text = "Transfer" Then
Frame1.Visible = True
End If
End Sub

Private Sub cmd_clr_Click()
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
ph_no.Text = ""
can_type.Text = ""
tr_from.Text = ""
tr_to.Text = ""
new_addr.Text = ""
tel_no.Text = ""
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
On Error Resume Next
rs.AddNew
rs.Fields(0) = cdt.Text
rs.Fields(1) = ctime.Text
rs.Fields(2) = cus_code.Text
rs.Fields(3) = cus_name.Text
rs.Fields(4) = addr.Text
rs.Fields(5) = ph_no.Text
rs.Fields(6) = can_type.Text
rs.Fields(7) = tr_from.Text
rs.Fields(8) = tr_to.Text
rs.Fields(9) = new_addr.Text
rs.Fields(10) = tel_no.Text
rs.Update
MsgBox "Saved"
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
ph_no.Text = ""
can_type.Text = ""
tr_from.Text = ""
tr_to.Text = ""
new_addr.Text = ""
tel_no.Text = ""

rs.Close
rs.Open "select * from dis_det", cn, adOpenDynamic, adLockPessimistic

End Sub

Private Sub cus_code_Click()
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
If cus_code.Text = rs1(1) Then
cus_name.Text = rs1(2)
addr.Text = rs1(3)
ph_no.Text = rs1(4)
End If
rs1.MoveNext
Wend
End If
End Sub

Private Sub Form_Load()

Frame1.Visible = False

cdt.Text = Format(Date, "dd/mm/yyyy")
ctime.Text = Format(Time, "HH:MM:SS")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from dis_det", cn, adOpenDynamic, adLockPessimistic

rs1.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
cus_code.AddItem (rs1(1))
'tr_from.AddItem (rs1(4))
'tr_to.AddItem (rs1(4))
rs1.MoveNext
Wend
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
rs1.Close
cn.Close
End Sub
