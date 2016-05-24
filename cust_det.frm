VERSION 5.00
Begin VB.Form frm_cus 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Details"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t1 
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
      Left            =   1200
      TabIndex        =   23
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   0
      Picture         =   "cust_det.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   3075
      TabIndex        =   22
      Top             =   480
      Width           =   3135
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3360
      TabIndex        =   16
      Top             =   6840
      Width           =   6015
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
         Left            =   4920
         Picture         =   "cust_det.frx":3167
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_del 
         Caption         =   "DELETE"
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
         Left            =   3720
         Picture         =   "cust_det.frx":3812
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1095
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
         Left            =   2520
         Picture         =   "cust_det.frx":3EBD
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1095
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
         Left            =   1320
         Picture         =   "cust_det.frx":4568
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1095
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
         Picture         =   "cust_det.frx":4C13
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6015
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   6015
      Begin VB.TextBox area_code 
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
         TabIndex        =   29
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox place 
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
         ItemData        =   "cust_det.frx":52BE
         Left            =   3000
         List            =   "cust_det.frx":52D1
         TabIndex        =   28
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox br_band 
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
         ItemData        =   "cust_det.frx":5314
         Left            =   3000
         List            =   "cust_det.frx":531E
         TabIndex        =   26
         Top             =   5400
         Width           =   2295
      End
      Begin VB.ComboBox gender 
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
         ItemData        =   "cust_det.frx":532B
         Left            =   3000
         List            =   "cust_det.frx":5335
         TabIndex        =   25
         Top             =   4200
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
         Left            =   3000
         TabIndex        =   8
         Top             =   960
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
         Left            =   3000
         TabIndex        =   7
         Top             =   360
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
         Height          =   525
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox cont_no 
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
         TabIndex        =   5
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox m_id 
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
         TabIndex        =   4
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Area Code"
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
         Left            =   720
         TabIndex        =   30
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Place"
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
         Left            =   720
         TabIndex        =   27
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Broad Band"
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
         Left            =   720
         TabIndex        =   15
         Top             =   5400
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
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   1935
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
         Left            =   720
         TabIndex        =   13
         Top             =   960
         Width           =   2055
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
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1680
         Width           =   2055
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
         Left            =   720
         TabIndex        =   11
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mail ID"
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
         Left            =   720
         TabIndex        =   10
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gender"
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
         Left            =   720
         TabIndex        =   9
         Top             =   4200
         Width           =   2055
      End
   End
   Begin VB.TextBox dt 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label10 
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
      Left            =   240
      TabIndex        =   24
      Top             =   6960
      Width           =   1215
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
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
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
      Width           =   9735
   End
End
Attribute VB_Name = "frm_cus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmd_clr_Click()

cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
place.Text = ""
area_code.Text = ""
cont_no.Text = ""
Gender.Text = ""
m_id.Text = ""
br_band.Text = ""

If rs.RecordCount = 0 Then
cus_code.Text = "2000"
Else
rs.MoveLast
cus_code.Text = rs(1) + 1
End If

End Sub

Private Sub cmd_del_Click()

If cus_code.Text = "" Or cus_name.Text = "" Or addr.Text = "" Or cont_no.Text = "" Or Gender.Text = "" Or m_id.Text = "" Or br_band.Text = "" Then
MsgBox "Empty Fields"
Else
cn.Execute "delete from cus_info where ccode='" + cus_code.Text + "'"
MsgBox "Deleted"
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
place.Text = ""
area_code.Text = ""
cont_no.Text = ""
Gender.Text = ""
m_id.Text = ""
br_band.Text = ""

rs.Close
rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic

cus_code.Clear
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
cus_code.AddItem (rs(1))
rs.MoveNext
Wend
End If
End If
End Sub

Private Sub cmd_edit_Click()
If cus_code.Text = "" Or cus_name.Text = "" Or addr.Text = "" Or cont_no.Text = "" Or Gender.Text = "" Or m_id.Text = "" Or br_band.Text = "" Then
MsgBox "Empty Fields"
Else
cn.Execute "update cus_info set cname='" + cus_name.Text + "',addr='" + addr.Text + "',cont_no='" + cont_no.Text + "',gender='" + Gender.Text + "',m_id='" + m_id.Text + "',br_band='" + br_band.Text + "',place='" + place.Text + "',area_code='" + area_code.Text + "' where ccode='" + cus_code.Text + "'"
MsgBox "Updated"
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
place.Text = ""
area_code.Text = ""
cont_no.Text = ""
Gender.Text = ""
m_id.Text = ""
br_band.Text = ""


rs.Close
rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
cus_code.AddItem (rs(1))
rs.MoveNext
Wend
End If

End If
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
'On Error Resume Next
If cus_code.Text = "" Or cus_name.Text = "" Or addr.Text = "" Or cont_no.Text = "" Or Gender.Text = "" Or m_id.Text = "" Or br_band.Text = "" Then
MsgBox "Empty Fields"
Else
rs.AddNew
rs.Fields(0) = dt.Text
rs.Fields(1) = cus_code.Text
rs.Fields(2) = cus_name.Text
rs.Fields(3) = addr.Text
rs.Fields(4) = place.Text
rs.Fields(5) = area_code.Text
rs.Fields(6) = cont_no.Text
rs.Fields(7) = Gender.Text
rs.Fields(8) = m_id.Text
rs.Fields(9) = br_band.Text
rs.Update
MsgBox "Saved"
cus_code.Text = ""
cus_name.Text = ""
addr.Text = ""
place.Text = ""
area_code.Text = ""
cont_no.Text = ""
Gender.Text = ""
m_id.Text = ""
br_band.Text = ""

rs.Close
rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic

cus_code.Clear
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
cus_code.AddItem (rs(1))
rs.MoveNext
Wend
End If


If rs.RecordCount = 0 Then
cus_code.Text = "2000"
Else
rs.MoveLast
cus_code.Text = rs(1) + 1
End If

End If

End Sub
Private Sub cus_code_Click()
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
If cus_code.Text = rs(1) Then
dt.Text = rs.Fields(0)
cus_code.Text = rs.Fields(1)
cus_name.Text = rs.Fields(2)
addr.Text = rs.Fields(3)
place.Text = rs.Fields(4)
area_code.Text = rs.Fields(5)
cont_no.Text = rs.Fields(6)
Gender.Text = rs.Fields(7)
m_id.Text = rs.Fields(8)
br_band.Text = rs.Fields(9)
End If
rs.MoveNext
Wend
End If

End Sub

Private Sub Form_Load()

dt.Text = Format(Date, "dd/mm/yyyy")
t1.Text = Format(Time, "HH:MM:SS")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
cus_code.AddItem (rs(1))
rs.MoveNext
Wend
End If

If rs.RecordCount = 0 Then
cus_code.Text = "2000"
Else
rs.MoveLast
cus_code.Text = rs(1) + 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
cn.Close
End Sub

Private Sub place_Click()

If place.Text = "HOPES COLLEGE" Then
area_code.Text = "A5699"
ElseIf place.Text = "CHITRA" Then
area_code.Text = "A6000"
ElseIf place.Text = "SINGANALOUR" Then
area_code.Text = "A7890"
ElseIf place.Text = "PEELAMEDU" Then
area_code.Text = "A6544"
ElseIf place.Text = "GANDHI MANAGAR" Then
area_code.Text = "A9877"
End If

End Sub
