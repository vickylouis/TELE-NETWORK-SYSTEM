VERSION 5.00
Begin VB.Form frm_comp_det 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Complaint Details"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form6"
   ScaleHeight     =   7410
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   0
      Picture         =   "mail.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   30
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   2400
      TabIndex        =   21
      Top             =   5760
      Width           =   6015
      Begin VB.CommandButton cmd_last 
         Caption         =   "LAST"
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
         Left            =   4560
         Picture         =   "mail.frx":0A1A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmd_next 
         Caption         =   "NEXT"
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
         Left            =   3120
         Picture         =   "mail.frx":10C5
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_pre 
         Caption         =   "PREVIOUS"
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
         Left            =   1680
         Picture         =   "mail.frx":1770
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_first 
         Caption         =   "FIRST"
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
         Left            =   120
         Picture         =   "mail.frx":1E1B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   840
         Width           =   1215
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
         Left            =   120
         Picture         =   "mail.frx":24C6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "EDIT"
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
         Left            =   1680
         Picture         =   "mail.frx":2B71
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1215
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
         Left            =   3120
         Picture         =   "mail.frx":321C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
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
         Left            =   4560
         Picture         =   "mail.frx":38C7
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox status 
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
      ItemData        =   "mail.frx":3F72
      Left            =   7680
      List            =   "mail.frx":3F7C
      TabIndex        =   20
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox comp_acc_dt 
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
      Left            =   2880
      TabIndex        =   17
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Problem Notification"
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
      Height          =   2295
      Left            =   600
      TabIndex        =   12
      Top             =   2400
      Width           =   4455
      Begin VB.TextBox comp_desc 
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
         Height          =   1005
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint Description"
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
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Details"
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
      Height          =   3735
      Left            =   5520
      TabIndex        =   5
      Top             =   960
      Width           =   4575
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
         Left            =   2160
         TabIndex        =   15
         Top             =   3120
         Width           =   2295
      End
      Begin VB.ComboBox cus_code 
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
         Left            =   2160
         TabIndex        =   11
         Top             =   480
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
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1920
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
         Left            =   2160
         TabIndex        =   6
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telephone No"
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
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer name"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Code"
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
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox comp_ref 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox comp_dt 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
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
      Left            =   5520
      TabIndex        =   19
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comp Access Date"
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
      Left            =   720
      TabIndex        =   18
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "COMPLAINT DETAILS"
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
      Width           =   10455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comp Date"
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
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comp Ref"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frm_comp_det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmd_clr_Click()

comp_ref.Text = ""
comp_desc.Text = ""
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
ph_no.Text = ""
comp_acc_dt.Text = ""
Status.Text = ""

If rs.RecordCount = 0 Then
Else
rs.MoveLast
comp_ref.Text = rs(0) + 1
End If

End Sub

Private Sub cmd_edit_Click()

If comp_ref.Text = "" Or comp_desc.Text = "" Or cus_code.Text = "" Or cus_name.Text = "" Or addr.Text = "" Or ph_no.Text = "" Or comp_acc_dt.Text = "" Or Status.Text = "" Then
MsgBox "Empty Fields"
Else
cn.Execute "update comp_details set comp_acc_dt='" + comp_acc_dt.Text + "',status='" + Status.Text + "' where comp_ref='" + comp_ref.Text + "'"
MsgBox "Updated"
comp_ref.Text = ""
comp_desc.Text = ""
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
ph_no.Text = ""
comp_acc_dt.Text = ""
Status.Text = ""
End If

rs.Close
rs.Open "select * from comp_details", cn, adOpenDynamic, adLockPessimistic

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_first_Click()
On Error Resume Next
rs.MoveFirst
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
End Sub

Private Sub cmd_last_Click()
On Error Resume Next
rs.MoveLast
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
End Sub

Private Sub cmd_next_Click()
On Error Resume Next
rs.MoveNext
If rs.EOF = True Then
MsgBox "Last Record"
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
  Else
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
End If
End Sub

Private Sub cmd_pre_Click()
On Error Resume Next
rs.MovePrevious
If rs.BOF = True Then
MsgBox "First Record"
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
  Else
  comp_ref.Text = rs.Fields(0)
  comp_dt.Text = rs.Fields(1)
  comp_desc.Text = rs.Fields(2)
  cus_code.Text = rs.Fields(3)
  cus_name.Text = rs.Fields(4)
  addr.Text = rs.Fields(5)
  ph_no.Text = rs.Fields(6)
  comp_acc_dt.Text = rs.Fields(7)
  Status.Text = rs.Fields(8)
End If
End Sub

Private Sub cmd_save_Click()

If comp_ref.Text = "" Or comp_desc.Text = "" Or cus_code.Text = "" Or cus_name.Text = "" Or addr.Text = "" Or ph_no.Text = "" Or comp_acc_dt.Text = "" Or Status.Text = "" Then
MsgBox "Empty Fields"
Else
rs.AddNew
rs.Fields(0) = comp_ref.Text
rs.Fields(1) = comp_dt.Text
rs.Fields(2) = comp_desc.Text
rs.Fields(3) = cus_code.Text
rs.Fields(4) = cus_name.Text
rs.Fields(5) = addr.Text
rs.Fields(6) = ph_no.Text
rs.Fields(7) = comp_acc_dt.Text
rs.Fields(8) = Status.Text
rs.Update
MsgBox "Saved"
comp_ref.Text = ""
comp_desc.Text = ""
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
ph_no.Text = ""
comp_acc_dt.Text = ""
Status.Text = ""
End If

rs.Close
rs.Open "select * from comp_details", cn, adOpenDynamic, adLockPessimistic

If rs.RecordCount = 0 Then
Else
rs.MoveLast
comp_ref.Text = rs(0) + 1
End If

End Sub

Private Sub cus_code_Click()
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
If cus_code.Text = rs1(1) Then
cus_name.Text = rs1(2)
addr.Text = rs1(3)
ph_no.Text = rs1(6)
End If
rs1.MoveNext
Wend
End If
End Sub

Private Sub Form_Load()

comp_dt.Text = Format(Date, "dd/mm/yy")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from comp_details", cn, adOpenDynamic, adLockPessimistic

rs1.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
cus_code.AddItem (rs1(1))
rs1.MoveNext
Wend
End If

If rs.RecordCount = 0 Then
comp_ref.Text = "2000"
Else
rs.MoveLast
comp_ref.Text = rs(0) + 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
rs1.Close
cn.Close
End Sub

Private Sub Picture2_Click()
Form1.Show
End Sub
