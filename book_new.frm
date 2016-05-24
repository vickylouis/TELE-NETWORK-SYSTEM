VERSION 5.00
Begin VB.Form frm_booking 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Booking of New Phone"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   7560
      TabIndex        =   15
      Top             =   1440
      Width           =   2295
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
         Left            =   480
         Picture         =   "book_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
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
         Left            =   480
         Picture         =   "book_new.frx":06AB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
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
         Left            =   480
         Picture         =   "book_new.frx":0D56
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Booking"
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
      Height          =   2655
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
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
         Left            =   2400
         TabIndex        =   10
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
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox dep_amt 
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
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox ph_type 
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
         ItemData        =   "book_new.frx":1401
         Left            =   2400
         List            =   "book_new.frx":140B
         TabIndex        =   7
         Top             =   2160
         Width           =   2295
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
         Left            =   120
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
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount of Deposit"
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
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone Type"
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
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
   End
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
      Left            =   8400
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   -360
      Picture         =   "book_new.frx":1425
      ScaleHeight     =   3795
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   480
      Width           =   2775
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
      Left            =   5400
      TabIndex        =   2
      Top             =   720
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
      Left            =   7440
      TabIndex        =   5
      Top             =   720
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
      Left            =   4440
      TabIndex        =   1
      Top             =   720
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
      Width           =   10215
   End
End
Attribute VB_Name = "frm_booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmd_clr_Click()
cus_code.Text = ""
cus_name.Text = ""
dep_amt.Text = ""
ph_type.Text = ""
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()

If cus_code.Text = "" Or cus_name.Text = "" Or dep_amt.Text = "" Or ph_type.Text = "" Then
MsgBox "Empty Fields"
Else
rs.AddNew
rs.Fields(0) = dt.Text
rs.Fields(1) = cus_code.Text
rs.Fields(2) = cus_name.Text
rs.Fields(3) = dep_amt.Text
rs.Fields(4) = ph_type.Text
rs.Update
MsgBox "Saved"
cus_code.Text = ""
cus_name.Text = ""
dep_amt.Text = ""
ph_type.Text = ""

rs.Close
rs.Open "select * from booking", cn, adOpenDynamic, adLockPessimistic
End If

End Sub

Private Sub cus_code_Click()
If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
While Not rs1.EOF
If cus_code.Text = rs1(1) Then
cus_name.Text = rs1(2)
End If
rs1.MoveNext
Wend
End If

End Sub

Private Sub Form_Load()

dt.Text = Format(Date, "dd/mm/yyyy")
t1.Text = Format(Time, "HH:MM:SS")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from booking", cn, adOpenDynamic, adLockPessimistic

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

End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
rs1.Close
cn.Close
End Sub
