VERSION 5.00
Begin VB.Form frm_reading 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reading Details"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form4"
   ScaleHeight     =   6090
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
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
      Left            =   2280
      Picture         =   "sales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5280
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
      Left            =   3840
      Picture         =   "sales.frx":06AB
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5280
      Width           =   1335
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
      Left            =   5520
      Picture         =   "sales.frx":0D56
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5280
      Width           =   1335
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
      Left            =   5160
      TabIndex        =   19
      Top             =   840
      Width           =   1575
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
      Left            =   7800
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox ph_no 
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
      Left            =   2400
      TabIndex        =   17
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox isd_charge 
      Alignment       =   2  'Center
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
      Left            =   6960
      TabIndex        =   16
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox isd_unit 
      Alignment       =   2  'Center
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
      Left            =   6960
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox std_charge 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox std_unit 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   13
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox loc_charge 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   12
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox loc_unit 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox cus_code 
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
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox tran_dt 
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
      Left            =   7080
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox tran_no 
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
      Left            =   7080
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label4 
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
      Left            =   4200
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   6840
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Local Unit / Charges"
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
      TabIndex        =   10
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STD Unit / Charges"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ISD Unit / Charges"
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
      Left            =   6960
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trans Date"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone no."
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
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trans Number"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   2055
      Left            =   360
      Top             =   2880
      Width           =   9135
   End
   Begin VB.Label Label8 
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
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
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
      Width           =   9975
   End
End
Attribute VB_Name = "frm_reading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmd_clr_Click()
ph_no.Text = ""
cus_code.Text = ""
tran_no.Text = ""
'tran_dt.Text = ""
loc_unit.Text = ""
loc_charge.Text = ""
std_unit.Text = ""
std_charge.Text = ""
isd_unit.Text = ""
isd_charge.Text = ""

If rs.RecordCount = 0 Then
Else
rs.MoveLast
tran_no.Text = rs(4) + 1
End If

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()

rs.AddNew
rs.Fields(0) = cdt.Text
rs.Fields(1) = ctime.Text
rs.Fields(2) = ph_no.Text
rs.Fields(3) = cus_code.Text
rs.Fields(4) = tran_no.Text
rs.Fields(5) = tran_dt.Text
rs.Fields(6) = loc_unit.Text
rs.Fields(7) = loc_charge.Text
rs.Fields(8) = std_unit.Text
rs.Fields(9) = std_charge.Text
rs.Fields(10) = isd_unit.Text
rs.Fields(11) = isd_charge.Text
rs.Update
MsgBox "Saved"
ph_no.Text = ""
cus_code.Text = ""
tran_no.Text = ""
'tran_dt.Text = ""
loc_unit.Text = ""
loc_charge.Text = ""
std_unit.Text = ""
std_charge.Text = ""
isd_unit.Text = ""
isd_charge.Text = ""

rs.Close
rs.Open "select * from reading", cn, adOpenDynamic, adLockPessimistic

If rs.RecordCount = 0 Then
Else
rs.MoveLast
tran_no.Text = rs(4) + 1
End If

End Sub

Private Sub Form_Load()

cdt.Text = Format(Date, "dd/mm/yyyy")
ctime.Text = Format(Time, "HH:MM:SS")
tran_dt.Text = Format(Date, "dd/mm/yyyy")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from reading", cn, adOpenDynamic, adLockPessimistic

rs1.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
ph_no.Clear
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
ph_no.AddItem (rs1(6))
rs1.MoveNext
Wend
End If

If rs.RecordCount = 0 Then
Else
rs.MoveLast
tran_no.Text = rs(4) + 1
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
rs.Close
rs1.Close
cn.Close
End Sub

Private Sub isd_unit_Change()
isd_charge.Text = Val(isd_unit) * 10
End Sub

Private Sub loc_unit_Change()
loc_charge.Text = Val(loc_unit) * 2
End Sub

Private Sub ph_no_Click()
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
If ph_no.Text = rs1(6) Then
cus_code.Text = rs1(2)
End If
rs1.MoveNext
Wend
End If
End Sub

Private Sub Text5_Change()

End Sub

Private Sub std_unit_Change()
std_charge.Text = Val(std_unit) * 4

End Sub
