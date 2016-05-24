VERSION 5.00
Begin VB.Form frm_bill 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bill Collection"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form3"
   ScaleHeight     =   7080
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox bill_no 
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
      Left            =   2640
      TabIndex        =   27
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmd_read 
      Caption         =   "READING"
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
      Left            =   6240
      Picture         =   "bill.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox ph_type 
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
      Left            =   2640
      TabIndex        =   25
      Top             =   4320
      Width           =   2295
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
      Left            =   4680
      Picture         =   "bill.frx":06AB
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
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
      Left            =   3120
      Picture         =   "bill.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6360
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
      Left            =   1560
      Picture         =   "bill.frx":1401
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   5280
      TabIndex        =   13
      Top             =   3000
      Width           =   4935
      Begin VB.TextBox bank_name 
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
         Height          =   645
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox chq_no 
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
         TabIndex        =   19
         Text            =   "0"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox pay_type 
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
         ItemData        =   "bill.frx":1AAC
         Left            =   2400
         List            =   "bill.frx":1AB6
         TabIndex        =   17
         Top             =   360
         Width           =   2295
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
         Left            =   2400
         TabIndex        =   15
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bank Name"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cheque No"
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
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payment Type"
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
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2055
      End
   End
   Begin VB.TextBox ex_code 
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
      Left            =   2640
      TabIndex        =   12
      Top             =   5520
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      Picture         =   "bill.frx":1AC8
      ScaleHeight     =   1515
      ScaleWidth      =   10395
      TabIndex        =   10
      Top             =   480
      Width           =   10455
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
      Left            =   2640
      TabIndex        =   9
      Top             =   4920
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
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
      Left            =   2640
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
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
      Height          =   285
      Left            =   8400
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bill Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exchange Code"
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
      TabIndex        =   11
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label7 
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
      Left            =   7440
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone Number"
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
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone Type"
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
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
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
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frm_bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub cmd_clr_Click()
cus_code.Text = ""
cus_name.Text = ""
ph_type.Text = ""
ph_no.Text = ""
ex_code.Text = ""
pay_type.Text = ""
chq_no.Text = ""
bank_name.Text = ""
amt.Text = ""

If rs.RecordCount = 0 Then
Else
rs.MoveLast
bill_no.Text = rs(1) + 1
End If

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_read_Click()
frm_reading.Show
End Sub

Private Sub cmd_save_Click()

If cus_code.Text = "" Or cus_name.Text = "" Or ph_type.Text = "" Or ph_no.Text = "" Or ex_code.Text = "" Or pay_type.Text = "" Or chq_no.Text = "" Or bank_name.Text = "" Or amt.Text = "" Then
MsgBox "Empty Fields"
Else
rs.AddNew
rs.Fields(0) = cdt.Text
rs.Fields(1) = bill_no.Text
rs.Fields(2) = cus_code.Text
rs.Fields(3) = cus_name.Text
rs.Fields(4) = ph_type.Text
rs.Fields(5) = ph_no.Text
rs.Fields(6) = ex_code.Text
rs.Fields(7) = pay_type.Text
rs.Fields(8) = chq_no.Text
rs.Fields(9) = bank_name.Text
rs.Fields(10) = amt.Text
rs.Update
MsgBox "Saved"
cus_code.Text = ""
cus_name.Text = ""
ph_type.Text = ""
ph_no.Text = ""
ex_code.Text = ""
pay_type.Text = ""
chq_no.Text = ""
bank_name.Text = ""
amt.Text = ""

rs.Close
rs.Open "select * from bill_details", cn, adOpenDynamic, adLockPessimistic

If rs.RecordCount = 0 Then
Else
rs.MoveLast
bill_no.Text = rs(1) + 1
End If

End If

End Sub

Private Sub cus_code_Click()

amt.Text = ""

If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
If cus_code.Text = rs1(1) Then
cus_name.Text = rs1(2)
ph_type.Text = rs1(4)
End If
rs1.MoveNext
Wend
End If

If rs2.RecordCount = 0 Then
Else
rs2.MoveFirst
While Not rs2.EOF
If cus_name.Text = rs2(3) Then
ph_no.Text = rs2(2)
amt.Text = Val(rs2(7)) + Val(rs2(9)) + Val(rs2(11))
End If
rs2.MoveNext
Wend
End If

End Sub

Private Sub Form_Load()

cdt.Text = Format(Date, "dd/mm/yyyy")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from bill_details", cn, adOpenDynamic, adLockPessimistic

rs1.Open "select * from booking", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
cus_code.AddItem (rs1(1))
rs1.MoveNext
Wend
End If

rs2.Open "select * from reading", cn, adOpenDynamic, adLockPessimistic

If rs.RecordCount = 0 Then
Else
rs.MoveLast
bill_no.Text = rs(1) + 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
rs1.Close
cn.Close
End Sub
